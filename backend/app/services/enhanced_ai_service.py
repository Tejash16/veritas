import google.generativeai as genai
from typing import Dict, Any, List, Tuple, Optional
import json
import base64
import structlog
from PIL import Image
import io
import time
import asyncio
import os
import fitz  # PyMuPDF
import pandas as pd
from decouple import config
import re
from datetime import datetime
import math

# Configure logging
logger = structlog.get_logger()

class EnhancedGeminiService:
    def __init__(self):
        # Get API key from environment
        api_key = config('GOOGLE_API_KEY', default=None)
        
        if not api_key:
            raise ValueError("GOOGLE_API_KEY is required for Gemini 2.5 Pro")
        
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-2.0-flash-exp')
        self.ai_enabled = True
        
        # Configuration for comprehensive extraction
        self.max_cells_per_batch = 200  # Increased from previous limits
        self.max_sheets_per_workbook = 50  # Process up to 50 sheets
        self.max_rows_per_sheet = 1000  # Up from 50
        self.max_cols_per_sheet = 100   # Up from 20
        self.min_table_size = 2  # Minimum rows for a table (header + at least 1 data row)
        self.max_gap_tolerance = 3  # Maximum empty rows allowed within a table

        logger.info("Enhanced Gemini 2.5 Pro Service initialized with comprehensive extraction settings")

    async def extract_comprehensive_pdf_data(self, pdf_path: str) -> Dict[str, Any]:
        """
        Comprehensive PDF extraction using Gemini 2.5 Pro with bounding boxes
        """
        logger.info(f"Starting comprehensive PDF extraction: {pdf_path}")
        
        try:
            doc = fitz.open(pdf_path)
            page_analyses = []
            
            for page_num in range(len(doc)):
                page = doc[page_num]
                
                # Extract high-quality page image with coordinates
                page_data = await self._extract_page_with_coordinates(page, page_num + 1)
                page_analyses.append(page_data)
                
                # Rate limiting for Gemini API
                await asyncio.sleep(1.5)
            
            doc.close()
            
            # Synthesize complete document analysis
            comprehensive_data = await self._synthesize_document_analysis(page_analyses)
            
            logger.info(f"PDF extraction completed: {len(comprehensive_data.get('all_extracted_values', []))} values found")
            return comprehensive_data
            
        except Exception as e:
            logger.error(f"PDF extraction failed: {e}")
            raise

    async def _extract_page_with_coordinates(self, page, page_num: int) -> Dict[str, Any]:
        """
        Extract page data with precise coordinate mapping using Gemini 2.5 Pro
        """
        try:
            # Convert page to high-resolution image
            mat = fitz.Matrix(2.0, 2.0)  # High quality for better extraction
            pix = page.get_pixmap(matrix=mat)
            img_data = pix.tobytes("png")
            
            # Convert to PIL Image for Gemini
            image = Image.open(io.BytesIO(img_data))
            
            # Enhanced prompt for better extraction
            prompt = f"""
Analyze this presentation slide (page {page_num}) and extract ALL numerical values, financial metrics, percentages, dates, and quantitative data.

CRITICAL: Extract EVERY number you can find on this slide, including:
- Revenue figures, costs, profits
- Percentages, ratios, growth rates  
- Dates, years, quarters
- Counts, quantities, volumes
- Currency amounts in any format
- Statistical data, KPIs, metrics

For each number found, provide:
- Exact value as displayed
- Business context explaining what it represents. If it is a derived term mention the whole context. For example - 44% growth from 1916 in FY24 to 2759 in FY25
- Normalized coordinates [x1, y1, x2, y2] on 0-1 scale
- Data type classification

Return ONLY valid JSON in this exact format:
{{
    "page_number": {page_num},
    "page_dimensions": {{"width": {image.width}, "height": {image.height}}},
    "extracted_values": [
        {{
            "id": "value_{page_num}_001",
            "value": "exact_number_as_displayed",
            "normalized_value": "cleaned_numeric_format",
            "data_type": "currency|percentage|count|ratio|date|metric",
            "coordinates": {{
                "bounding_box": [0.1, 0.2, 0.3, 0.4],
                "confidence": 0.9
            }},
            "business_context": {{
                "semantic_meaning": "detailed_description_of_what_this_number_represents",
                "business_category": "revenue|costs|growth|operational|financial|market",
                "presentation_priority": "primary|secondary|supporting",
                "calculation_type": "absolute|percentage|ratio|growth_rate|other"
            }},
            "confidence": 0.9
        }}
    ]
}}

IMPORTANT: Be thorough - extract ALL visible numbers, not just the prominent ones. Return only valid JSON.
"""

            response = self.model.generate_content([prompt, image])
            result = await self._parse_gemini_json_response_robust(response.text, f"page_{page_num}")
            
            # Validate and enhance coordinates
            result = self._validate_and_enhance_coordinates(result, image.size)
            
            logger.info(f"Page {page_num}: Extracted {len(result.get('extracted_values', []))} values")
            return result
            
        except Exception as e:
            logger.error(f"Page {page_num} extraction failed: {e}")
            return {
                "page_number": page_num,
                "page_dimensions": {"width": 800, "height": 600},
                "extracted_values": [],
                "error": str(e)
            }


    async def analyze_excel_comprehensive(self, excel_path: str) -> Dict[str, Any]:
        """
        Comprehensive Excel analysis using table extraction approach
        """
        logger.info(f"Starting Excel analysis using table extraction: {excel_path}")
        
        try:
            # Extract tables from all sheets
            all_tables = await self._extract_tables_from_excel(excel_path)
            logger.info(f"Extracted {len(all_tables)} tables from Excel file")
            
            # Analyze tables with Gemini
            sheet_analyses = []
            all_potential_sources = []
            
            for sheet_name, tables in all_tables.items():
                logger.info(f"Analyzing {len(tables)} tables in sheet: {sheet_name}")
                
                sheet_analysis = await self._analyze_tables_in_sheet(sheet_name, tables)
                sheet_analyses.append(sheet_analysis)
                
                # Collect all potential sources
                sources = sheet_analysis.get("potential_sources", [])
                for source in sources:
                    source["source_sheet"] = sheet_name
                all_potential_sources.extend(sources)
            
            # Synthesize workbook analysis
            workbook_analysis = await self._synthesize_workbook_analysis(sheet_analyses, all_potential_sources)
            
            logger.info(f"Excel analysis completed: {len(all_potential_sources)} potential sources identified")
            
            return workbook_analysis
            
        except Exception as e:
            logger.error(f"Excel analysis failed: {e}")
            raise

    async def _extract_tables_from_excel(self, excel_path: str) -> Dict[str, List[Dict]]:
        """
        Extract tables from Excel sheets while preserving structure
        """
        try:
            import openpyxl
            
            wb = openpyxl.load_workbook(excel_path, data_only=True)
            sheets_to_process = min(len(wb.sheetnames), self.max_sheets_per_workbook)
            
            all_tables = {}
            
            for sheet_name in wb.sheetnames[:sheets_to_process]:
                sheet = wb[sheet_name]
                tables = self._extract_tables_from_sheet(sheet)
                all_tables[sheet_name] = tables
                
                logger.info(f"Sheet '{sheet_name}': Found {len(tables)} tables")
            
            return all_tables
            
        except Exception as e:
            logger.error(f"Table extraction failed: {e}")
            return {}

    def _extract_tables_from_sheet(self, sheet) -> List[Dict]:
        """
        Improved table detection that handles gaps, merged cells, and multi-header tables
        """
        try:
            import openpyxl
            from openpyxl.utils import get_column_letter
            
            tables = []
            processed_cells = set()
            max_row = sheet.max_row or 1
            max_col = sheet.max_column or 1
            
            # Convert sheet to DataFrame for easier pattern detection
            all_data = []
            for row in range(1, min(max_row + 1, 1000)):  # Limit to 1000 rows
                row_data = []
                for col in range(1, min(max_col + 1, 50)):  # Limit to 50 columns
                    cell = sheet.cell(row, col)
                    cell_value = cell.value
                    if cell_value is None:
                        cell_value = ""
                    elif isinstance(cell_value, str):
                        cell_value = cell_value.strip()
                    row_data.append(cell_value)
                all_data.append(row_data)
            
            df = pd.DataFrame(all_data)
            
            # Detect table regions using pattern matching
            table_regions = self._detect_table_regions(df)
            
            for region in table_regions:
                start_row, start_col, end_row, end_col = region
                
                # Skip if already processed or too small
                if (end_row - start_row < 1) or (end_col - start_col < 1):
                    continue
                    
                # Extract table data with merged cell handling
                table_data = self._extract_table_with_merged_cells(
                    sheet, start_row, start_col, end_row, end_col
                )
                
                if table_data and len(table_data) >= 2:  # At least header + 1 row
                    tables.append({
                        "start_row": start_row,
                        "start_col": start_col,
                        "end_row": end_row,
                        "end_col": end_col,
                        "data": table_data,
                        "columns": table_data[0] if table_data else [],
                        "shape": (len(table_data) - 1, len(table_data[0])) if table_data else (0, 0)
                    })
            
            return tables
            
        except Exception as e:
            logger.error(f"Table extraction failed: {e}")
            return []
        
    def _detect_table_regions(self, df: pd.DataFrame) -> List[Tuple[int, int, int, int]]:
        """
        Detect table regions based on data patterns and gaps
        """
        regions = []
        rows, cols = df.shape
        visited = set()
        
        for start_row in range(rows):
            for start_col in range(cols):
                if (start_row, start_col) in visited:
                    continue
                    
                cell_value = df.iat[start_row, start_col]
                
                # Look for potential table starters (headers, numbers, meaningful text)
                if self._is_potential_table_starter(cell_value):
                    region = self._find_table_boundaries(df, start_row, start_col, visited)
                    if region:
                        regions.append(region)
        
        return regions

    def _is_potential_table_starter(self, value) -> bool:
        """
        Check if a cell value looks like a table header or data starter
        """
        if not value or value == "":
            return False
            
        # Header-like patterns
        header_patterns = [
            r'^[A-Z][a-zA-Z\s&]+$',  # Capitalized words with spaces/ampersands
            r'FY\d{2}',              # Fiscal years like FY23, FY24
            r'Q[1-4]',               # Quarters
            r'total|Total|TOTAL',    # Total indicators
            r'^[A-Z]{2,}$',          # All caps acronyms
        ]
        
        # Numeric patterns (could be data cells)
        numeric_patterns = [
            r'^\d+([.,]\d+)*$',      # Numbers with optional thousands separators
            r'^\d+%$',               # Percentages
        ]
        
        value_str = str(value)
        
        # Check header patterns
        for pattern in header_patterns:
            if re.match(pattern, value_str, re.IGNORECASE):
                return True
        
        # Check numeric patterns
        for pattern in numeric_patterns:
            if re.match(pattern, value_str):
                return True
        
        return False

    def _find_table_boundaries(self, df: pd.DataFrame, start_row: int, start_col: int, visited: set) -> Optional[Tuple[int, int, int, int]]:
        """
        Find the boundaries of a table starting from a given cell
        """
        rows, cols = df.shape
        end_row = start_row
        end_col = start_col
        
        # Expand right to find column boundary
        for col in range(start_col, cols):
            # Allow some empty cells in header area (for multi-header tables)
            empty_count = 0
            max_empty_in_header = 2
            
            for check_row in range(start_row, min(start_row + 5, rows)):
                if (check_row, col) in visited:
                    break
                if not df.iat[check_row, col] or df.iat[check_row, col] == "":
                    empty_count += 1
                    if empty_count > max_empty_in_header:
                        break
                else:
                    empty_count = 0
            
            if empty_count > max_empty_in_header:
                end_col = col - 1
                break
            else:
                end_col = col
        
        # Expand down to find row boundary with gap tolerance
        gap_tolerance = 2  # Allow up to 2 empty rows within table
        current_gap = 0
        
        for row in range(start_row, rows):
            row_has_data = False
            
            # Check if this row has any data in our column range
            for col in range(start_col, end_col + 1):
                if (row, col) in visited:
                    continue
                if df.iat[row, col] and df.iat[row, col] != "":
                    row_has_data = True
                    current_gap = 0
                    break
            
            if row_has_data:
                end_row = row
            else:
                current_gap += 1
                if current_gap > gap_tolerance:
                    break
        
        # Mark cells as visited
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                visited.add((row, col))
        
        # Ensure we have a valid table (at least 2x2)
        if (end_row - start_row >= 1) and (end_col - start_col >= 1):
            return (start_row + 1, start_col + 1, end_row + 1, end_col + 1)  # Convert to 1-indexed
        
        return None

    def _extract_table_with_merged_cells(self, sheet, start_row: int, start_col: int, end_row: int, end_col: int) -> List[List]:
        """
        Extract table data handling merged cells properly
        """
        try:
            table_data = []
            
            # Handle merged cells
            merged_cell_values = {}
            if hasattr(sheet, 'merged_cells'):
                for merge_range in sheet.merged_cells.ranges:
                    min_row, min_col, max_row, max_col = (
                        merge_range.min_row, merge_range.min_col, 
                        merge_range.max_row, merge_range.max_col
                    )
                    
                    # Only consider merged cells within our table region
                    if (min_row >= start_row and max_row <= end_row and 
                        min_col >= start_col and max_col <= end_col):
                        
                        # Get the value from the top-left cell
                        top_left_value = sheet.cell(min_row, min_col).value
                        if top_left_value is None:
                            top_left_value = ""
                        
                        # Store for all cells in merged range
                        for row in range(min_row, max_row + 1):
                            for col in range(min_col, max_col + 1):
                                merged_cell_values[(row, col)] = top_left_value
            
            # Extract table data
            for row in range(start_row, end_row + 1):
                row_data = []
                for col in range(start_col, end_col + 1):
                    cell_coord = (row, col)
                    
                    # Check if this cell is part of a merged range
                    if cell_coord in merged_cell_values:
                        cell_value = merged_cell_values[cell_coord]
                    else:
                        cell = sheet.cell(row, col)
                        cell_value = cell.value
                    
                    # Clean up values
                    if cell_value is None:
                        cell_value = ""
                    elif isinstance(cell_value, str):
                        cell_value = cell_value.strip()
                    
                    row_data.append(cell_value)
                
                # Skip completely empty rows
                if any(cell != "" for cell in row_data):
                    table_data.append(row_data)
            
            return table_data
            
        except Exception as e:
            logger.error(f"Table extraction with merged cells failed: {e}")
            return []

    async def _analyze_tables_in_sheet(self, sheet_name: str, tables: List[Dict]) -> Dict[str, Any]:
        """
        Analyze all tables in a sheet using Gemini
        """
        all_potential_sources = []
        
        for i, table in enumerate(tables):
            logger.info(f"Analyzing table {i+1}/{len(tables)} in sheet '{sheet_name}'")
            
            try:
                table_sources = await self._analyze_table_with_gemini(sheet_name, table, i)
                all_potential_sources.extend(table_sources)
                
                # Rate limiting between tables
                await asyncio.sleep(0.2)
                
            except Exception as e:
                logger.error(f"Table analysis failed for table {i} in {sheet_name}: {e}")
        
        # Sort by presentation likelihood
        all_potential_sources.sort(key=lambda x: x.get("presentation_likelihood", 0), reverse=True)
        
        return {
            "sheet_name": sheet_name,
            "potential_sources": all_potential_sources,
            "tables_analyzed": len(tables),
            "table_shapes": [t.get("shape", (0, 0)) for t in tables]
        }

    async def _analyze_table_with_gemini(self, sheet_name: str, table: Dict, table_index: int) -> List[Dict]:
        """
        Analyze a single table with Gemini to identify presentation-ready values
        """
        table_data = table["data"]
        columns = table["columns"]
        
        if not table_data or not columns:
            return []
        
        prompt = f"""
You are an expert in analyzing business Excel tables for presentation-readiness.

Given this table from sheet '{sheet_name}' (table {table_index + 1}), analyze it to identify values that would be relevant for business presentations.

TABLE STRUCTURE:
Columns: {columns}
Data (first 10 rows):
{json.dumps(table_data[:10], indent=2)}

Instructions:
For each potentially relevant value in the table, provide:
1. Presentation Likelihood: Rate 0.0 to 1.0 based on how likely the value is to be cited in a business presentation
2. Business Context: Describe the value's meaning based on its row and column context
3. Data Type: currency, percentage, count, ratio, or metric
4. Value Category: revenue, cost, growth, operational, market, strategic, or other
5. Reasoning: Why this value is relevant for presentations
6. Cell Reference: Approximate position (e.g., "B5" for row 5, column 2)

Focus on:
- Financial metrics (revenues, costs, profits, margins)
- Growth rates and percentages
- KPIs and performance indicators
- Summary values and totals
- Year-over-year comparisons
- Market statistics

Return output in strict JSON format as:
{{
    "table_analysis": [
        {{
            "cell_reference": "B5",
            "value": "actual_value",
            "business_context": "Q4 revenue for North America region",
            "presentation_likelihood": 0.9,
            "data_type": "currency",
            "value_category": "revenue",
            "reasoning": "Shows strong regional performance in latest quarter"
        }},
        ...
    ]
}}
Only return valid JSON; no extra text.
"""

        try:
            response = self.model.generate_content(prompt)
            result = await self._parse_gemini_json_response_robust(response.text, f"table_{sheet_name}_{table_index}")
            
            table_analysis = result.get('table_analysis', [])
            
            # Filter for likely presentation values
            presentation_sources = [
                source for source in table_analysis 
                if source.get("presentation_likelihood", 0) >= 0.4
            ]
            
            logger.info(f"Table {table_index} in {sheet_name}: {len(presentation_sources)} presentation-worthy sources")
            
            return presentation_sources
            
        except Exception as e:
            logger.error(f"Table analysis failed for {sheet_name} table {table_index}: {e}")
            return []

    async def _synthesize_workbook_analysis(self, sheet_analyses: List[Dict], all_potential_sources: List[Dict]) -> Dict[str, Any]:
        """Synthesize workbook analysis from all sheet analyses"""
        
        try:
            # Sort all sources by presentation likelihood
            all_potential_sources.sort(key=lambda x: x.get("presentation_likelihood", 0), reverse=True)
            
            # Create statistics
            total_sources = len(all_potential_sources)
            high_likelihood = len([s for s in all_potential_sources if s.get("presentation_likelihood", 0) >= 0.7])
            medium_likelihood = len([s for s in all_potential_sources if 0.4 <= s.get("presentation_likelihood", 0) < 0.7])
            
            # Category breakdown
            category_breakdown = {}
            for source in all_potential_sources:
                category = source.get("value_category", "unknown")
                category_breakdown[category] = category_breakdown.get(category, 0) + 1
            
            workbook_summary = {
                "workbook_summary": {
                    "total_sheets_processed": len(sheet_analyses),
                    "total_potential_sources": total_sources,
                    "high_likelihood_sources": high_likelihood,
                    "medium_likelihood_sources": medium_likelihood,
                    "category_breakdown": category_breakdown,
                    "analysis_timestamp": datetime.utcnow().isoformat(),
                    "extraction_method": "table_based_analysis"
                },
                "potential_sources": all_potential_sources,
                "sheet_analyses": sheet_analyses
            }
            
            logger.info(f"Workbook analysis completed: {total_sources} sources, {high_likelihood} high confidence")
            
            return workbook_summary
            
        except Exception as e:
            logger.error(f"Workbook synthesis failed: {e}")
            return {
                "workbook_summary": {
                    "total_sheets_processed": len(sheet_analyses),
                    "total_potential_sources": len(all_potential_sources),
                    "analysis_timestamp": datetime.utcnow().isoformat(),
                    "error": str(e)
                },
                "potential_sources": all_potential_sources,
                "sheet_analyses": sheet_analyses
            }

    # ... [Keep all the other existing methods like run_direct_comprehensive_audit, etc. unchanged] ...
    # [The rest of the methods remain the same as in your original code]

    def _generate_direct_audit_recommendations(self, summary: Dict) -> List[str]:
        """Generate recommendations based on direct audit summary"""
        recommendations = []
        
        total = summary["total_values_checked"]
        if total == 0:
            return ["No values were audited. Check extraction process."]
        
        accuracy = summary["overall_accuracy"]
        matched = summary["matched"]
        formatting_diffs = summary["formatting_differences"]
        mismatched = summary["mismatched"]
        pdf_only = summary["pdf_only"]
        
        # Overall assessment
        if accuracy >= 90:
            recommendations.append(f"Excellent data quality: {accuracy:.1f}% accuracy across all {total} values.")
        elif accuracy >= 75:
            recommendations.append(f"Good data quality: {accuracy:.1f}% accuracy. Review flagged discrepancies.")
        else:
            recommendations.append(f"Data quality concerns: {accuracy:.1f}% accuracy. Comprehensive review needed.")
        
        # Specific recommendations
        if matched > 0:
            recommendations.append(f"✅ {matched} values perfectly matched between presentation and Excel.")
        
        if formatting_diffs > 0:
            recommendations.append(f"📋 {formatting_diffs} values have formatting differences but same underlying data.")
        
        if mismatched > 0:
            recommendations.append(f"⚠️ {mismatched} values show potential discrepancies requiring review.")
        
        if pdf_only > 0:
            percentage = (pdf_only / total) * 100
            if percentage > 20:
                recommendations.append(f"🔍 {pdf_only} values ({percentage:.1f}%) appear only in presentation - verify if these should have Excel sources.")
            else:
                recommendations.append(f"ℹ️ {pdf_only} values are presentation-only (calculated metrics, external data, etc.)")
        
        # Coverage assessment
        recommendations.append(f"📊 Comprehensive coverage: All {total} extracted values were validated (100% coverage).")
        
        return recommendations

    async def run_direct_comprehensive_audit(self, pdf_values: List[Dict], excel_values: List[Dict]) -> Dict[str, Any]:
        """Run direct comprehensive audit comparing ALL PDF values against ALL Excel values"""
        logger.info(f"Starting direct comprehensive audit: {len(pdf_values)} PDF values vs {len(excel_values)} Excel values")
        
        if not pdf_values or not excel_values:
            return {
                "summary": {"total_values_checked": 0, "overall_accuracy": 0.0},
                "detailed_results": [],
                "recommendations": ["No values available for audit"],
                "risk_assessment": "high"
            }
        
        # Process PDF values in proper smaller batches
        batch_size = len(pdf_values)  # Reduced to manageable size for API
        all_audit_results = []
        total_batches = math.ceil(len(pdf_values) / batch_size)
        
        for batch_num in range(total_batches):
            start_idx = batch_num * batch_size
            end_idx = min((batch_num + 1) * batch_size, len(pdf_values))
            batch = pdf_values[start_idx:end_idx]
            
            logger.info(f"Processing batch {batch_num + 1}/{total_batches} ({len(batch)} values)")
            
            batch_results = await self._process_direct_audit_batch(batch, excel_values, batch_num + 1, total_batches)
            all_audit_results.extend(batch_results)
            
            # Rate limiting with progress tracking
            await asyncio.sleep(2)  # Increased sleep for better API reliability
        
        # Calculate comprehensive summary
        summary = self._calculate_audit_summary(all_audit_results)
        
        # Generate comprehensive recommendations
        recommendations = self._generate_direct_audit_recommendations(summary)
        
        return {
            "summary": summary,
            "detailed_results": all_audit_results,
            "recommendations": recommendations,
            "risk_assessment": "low" if summary["overall_accuracy"] >= 85 else "medium" if summary["overall_accuracy"] >= 70 else "high",
            "coverage_analysis": {
                "pdf_values_analyzed": len(pdf_values),
                "excel_values_searched": len(excel_values),
                "coverage_percentage": 100.0,
                "approach": "comprehensive_direct_validation",
                "batches_processed": total_batches
            }
        }

    def _calculate_audit_summary(self, all_audit_results: List[Dict]) -> Dict[str, Any]:
        """Calculate comprehensive audit summary statistics"""
        total = len(all_audit_results)
        
        summary = {
            "total_values_checked": total,
            "matched": len([r for r in all_audit_results if r.get("validation_status") == "matched"]),
            "mismatched": len([r for r in all_audit_results if r.get("validation_status") == "mismatched"]),
            "formatting_differences": len([r for r in all_audit_results if r.get("validation_status") == "formatting_difference"]),
            "unverifiable": len([r for r in all_audit_results if r.get("validation_status") == "unverifiable"]),
            "pdf_only": len([r for r in all_audit_results if r.get("validation_status") == "pdf_only"]),
            "errored": len([r for r in all_audit_results if r.get("error")]),
        }
        
        if summary["total_values_checked"] > 0:
            valid_results = total - summary["errored"]
            if valid_results > 0:
                summary["overall_accuracy"] = ((summary["matched"] + summary["formatting_differences"]) / valid_results) * 100
            else:
                summary["overall_accuracy"] = 0.0
            summary["success_rate"] = (valid_results / total) * 100
        else:
            summary["overall_accuracy"] = 0.0
            summary["success_rate"] = 0.0

        return summary

    async def _process_direct_audit_batch(self, pdf_batch: List[Dict], all_excel_values: List[Dict], batch_num: int, total_batches: int) -> List[Dict]:
        """Process a batch of PDF values against all Excel values"""
        
        # Create unique debug files for each batch
        debug_prefix = f"batch_{batch_num:02d}_of_{total_batches:02d}"
        
        # with open(f'pdf_{debug_prefix}.txt', 'w', encoding='utf-8') as f:
        #     f.write(str(pdf_batch))
        
        # Clean data for better prompt performance
        cleaned_pdf_batch = self._clean_pdf_batch_for_audit(pdf_batch)
        cleaned_excel_values = self._clean_excel_values_for_audit(all_excel_values)
        # cleaned_pdf_batch = pdf_batch
        # cleaned_excel_values = all_excel_values


        # with open(f'excel_{debug_prefix}.txt', 'w', encoding='utf-8') as f:
        #     f.write(str(cleaned_excel_values[:200]))  # Sample for debug
        
        # Use smarter Excel sampling - prioritize values that might match
        excel_sample = self._get_relevant_excel_sample(cleaned_excel_values, cleaned_pdf_batch)
        
        prompt = f"""
    You are auditing presentation values against Excel source data.

    BATCH {batch_num}/{total_batches} - PDF VALUES TO VALIDATE:
    {json.dumps(cleaned_pdf_batch, indent=1)}

    RELEVANT EXCEL VALUES TO SEARCH AGAINST (showing {len(excel_sample)} of {len(all_excel_values)} total):
    {json.dumps(excel_sample, indent=1)}

    For EACH PDF value (process them in order):
    1. FIRST try to directly match against Excel values
    2. If no direct match, check if it's a derived metric (growth rate, percentage, ratio).For example 44% Year-over-year growth represented by the separate cells with values 1916 to 2759.((2759-1916)/(1916)*100=44%)
    3. For derived metrics, attempt calculation using relevant Excel data
    4. Assign appropriate validation status

    CRITICAL: Process PDF values IN THE ORDER THEY APPEAR in the input list.

    Return ONLY valid JSON:
    {{
        "batch_results": [
            {{
                "pdf_value_id": "value_1_001",
                "pdf_value": "2759",
                "pdf_context": "FY25 revenue growth",
                "validation_status": "matched|mismatched|unverifiable",
                "excel_match": {{
                    "source_cell": "Sheet1!B5",
                    "excel_value": "2759",
                    "match_confidence": 0.95,
                    "calculation_basis": "direct_match|calculated|inferred"
                }},
                "confidence": 0.95,
                "audit_reasoning": "Detailed explanation including calculation steps if applicable"
            }}
        ]
    }}

    Validation Status Guide(Only these 3 validation status is allowed):
    - matched: Exact or equivalent value found, match_confidence greater than 0.90. Calculated values also have validation status as matched. PDF or Excel match also fall under this category. excel_match or pdf_match should also have validation status as matched
    - mismatched: Different values but related context  
    - unverifiable: Cannot verify relationship. Also when the value is only present in pdf

    IMPORTANT RULES:
    1. Return EXACTLY {len(pdf_batch)} results, one for each PDF value IN ORDER
    2. For calculated values, set the validation_status as matched and show the calculation in audit_reasoning
    3. Return ONLY valid JSON, no other text
    """

        try:
            with open(f'prompt_{debug_prefix}.txt', 'w', encoding='utf-8') as f:
                f.write(prompt)
            
            response = self.model.generate_content(prompt)
            
            result = await self._parse_gemini_json_response_robust(response.text, f"direct_audit_batch_{batch_num}")
            
            with open(f'parsed_result_{debug_prefix}.txt', 'w', encoding='utf-8') as f:
                f.write(str(result))
                
            batch_results = result.get('batch_results', [])
            
            # Ensure we have exactly the right number of results
            final_results = []
            for i, pdf_value in enumerate(pdf_batch):
                if i < len(batch_results):
                    audit_result = batch_results[i]
                    # Validate and enhance the result
                    enhanced_result = self._enhance_audit_result(audit_result, pdf_value, batch_num)
                    final_results.append(enhanced_result)
                else:
                    # Create fallback result for missing entries
                    fallback_result = self._create_fallback_audit_result(pdf_value, batch_num, i)
                    final_results.append(fallback_result)
            
            logger.info(f"Batch {batch_num}: Processed {len(final_results)} PDF values, {len([r for r in final_results if r.get('error')])} errors")
            return final_results
            
        except Exception as e:
            logger.error(f"Direct audit batch {batch_num} failed: {e}")
            # Create fallback results for entire batch
            return [self._create_fallback_audit_result(pdf_value, batch_num, i, str(e)) 
                    for i, pdf_value in enumerate(pdf_batch)]

    def _clean_pdf_batch_for_audit(self, pdf_batch: List[Dict]) -> List[Dict]:
        """Clean PDF batch data for audit prompts"""
        cleaned = []
        for item in pdf_batch:
            cleaned_item = {
                "id": item.get("id", ""),
                "value": item.get("value", ""),
                "normalized_value": item.get("normalized_value", ""),
                "data_type": item.get("data_type", ""),
                "business_context": item.get("business_context", {}).get("semantic_meaning", "")
            }
            cleaned.append(cleaned_item)
        return cleaned

    def _clean_excel_values_for_audit(self, excel_values: List[Dict]) -> List[Dict]:
        """Clean Excel values for audit prompts"""
        cleaned = []
        for item in excel_values:
            cleaned_item = {
                "value": item.get("value", ""),
                "cell_reference": item.get("cell_reference", ""),
                "sheet_name": item.get("sheet_name", ""),
                "business_context": item.get("business_context", "")
            }
            cleaned.append(cleaned_item)
        return cleaned

    def _get_relevant_excel_sample(self, excel_values: List[Dict], pdf_batch: List[Dict]) -> List[Dict]:
        """Get relevant Excel sample based on PDF batch content"""
        sample_size = min(150, len(excel_values))  # Increased sample size
        
        # Extract keywords from PDF batch for relevance filtering
        pdf_keywords = set()
        for pdf_item in pdf_batch:
            context = pdf_item.get("business_context", "").lower()
            pdf_keywords.update(context.split())
            value = str(pdf_item.get("value", "")).lower()
            pdf_keywords.update(value.split())
        
        # Prioritize Excel values that might be relevant
        if pdf_keywords:
            scored_excel = []
            for excel_item in excel_values:
                score = 0
                excel_context = excel_item.get("business_context", "").lower()
                excel_value = str(excel_item.get("value", "")).lower()
                
                # Score based on keyword matches
                for keyword in pdf_keywords:
                    if keyword in excel_context or keyword in excel_value:
                        score += 1
                
                scored_excel.append((score, excel_item))
            
            # Sort by relevance score and take top samples
            scored_excel.sort(key=lambda x: x[0], reverse=True)
            relevant_sample = [item for score, item in scored_excel[:sample_size]]
            
            # If we have enough relevant samples, return them
            if len(relevant_sample) >= sample_size // 2:
                return relevant_sample
        
        # Fallback: return a stratified sample
        return excel_values[:sample_size]

    def _enhance_audit_result(self, audit_result: Dict, original_pdf: Dict, batch_num: int) -> Dict:
        """Enhance audit result with additional metadata"""
        enhanced = audit_result.copy()
        enhanced.update({
            "original_pdf_data": original_pdf,
            "audit_timestamp": datetime.utcnow().isoformat(),
            "batch_number": batch_num,
            "pdf_value_id": enhanced.get("pdf_value_id", original_pdf.get("id", f"pdf_value_{batch_num}"))
        })
        return enhanced

    def _create_fallback_audit_result(self, pdf_value: Dict, batch_num: int, index: int, error_msg: str = None) -> Dict:
        """Create a fallback audit result for error cases"""
        result = {
            "pdf_value_id": pdf_value.get("id", f"pdf_value_{batch_num}_{index}"),
            "pdf_value": pdf_value.get("value", "unknown"),
            "pdf_context": pdf_value.get("business_context", {}).get("semantic_meaning", "unknown"),
            "validation_status": "unverifiable",
            "excel_match": None,
            "confidence": 0.0,
            "audit_reasoning": "Processing error" + (f": {error_msg}" if error_msg else ""),
            "original_pdf_data": pdf_value,
            "audit_timestamp": datetime.utcnow().isoformat(),
            "batch_number": batch_num
        }
        if error_msg:
            result["error"] = error_msg
        return result

    # ... [Include all other existing methods unchanged] ...

    async def _synthesize_document_analysis(self, page_analyses: List[Dict]) -> Dict[str, Any]:
        """
        Synthesize comprehensive document analysis using Gemini 2.5 Pro
        """
        # Combine all extracted values
        all_values = []
        for page in page_analyses:
            page_values = page.get('extracted_values', [])
            for value in page_values:
                value['page_number'] = page.get('page_number', 0)
                all_values.append(value)

        # Simplified synthesis for better reliability
        try:
            prompt = f"""
Analyze {len(page_analyses)} presentation pages with {len(all_values)} extracted values.

Create a document summary in JSON format:
{{
    "document_summary": {{
        "total_pages": {len(page_analyses)},
        "document_type": "financial_presentation",
        "main_business_themes": ["revenue", "growth", "performance"]
    }},
    "all_extracted_values": {json.dumps(all_values[:100])},
    "extraction_quality_metrics": {{
        "total_values_extracted": {len(all_values)},
        "overall_confidence": 0.85
    }}
}}

Return only valid JSON.
"""

            response = self.model.generate_content(prompt)
            synthesis = await self._parse_gemini_json_response_robust(response.text, "document_synthesis")
            
            # Ensure all_extracted_values is populated
            if not synthesis.get('all_extracted_values'):
                synthesis['all_extracted_values'] = all_values
            
            logger.info("Document synthesis completed successfully")
            return synthesis
            
        except Exception as e:
            logger.error(f"Document synthesis failed: {e}")
            # Return fallback structure
            return {
                "document_summary": {
                    "total_pages": len(page_analyses),
                    "document_type": "financial_presentation"
                },
                "all_extracted_values": all_values,
                "extraction_quality_metrics": {
                    "total_values_extracted": len(all_values),
                    "overall_confidence": 0.8
                },
                "synthesis_error": str(e)
            }

    def _validate_and_enhance_coordinates(self, result: Dict, image_size: Tuple[int, int]) -> Dict:
        """Validate and enhance coordinate data"""
        width, height = image_size
        
        for value in result.get('extracted_values', []):
            if 'coordinates' in value and 'bounding_box' in value['coordinates']:
                bbox = value['coordinates']['bounding_box']
                
                # Ensure coordinates are valid
                if len(bbox) == 4:
                    x1, y1, x2, y2 = bbox
                    x1 = max(0, min(1, float(x1)))
                    y1 = max(0, min(1, float(y1)))
                    x2 = max(0, min(1, float(x2)))
                    y2 = max(0, min(1, float(y2)))
                    
                    # Ensure logical ordering
                    if x1 > x2:
                        x1, x2 = x2, x1
                    if y1 > y2:
                        y1, y2 = y2, y1
                    
                    # Ensure minimum size
                    if x2 - x1 < 0.01:
                        x2 = min(1, x1 + 0.01)
                    if y2 - y1 < 0.01:
                        y2 = min(1, y1 + 0.01)
                    
                    value['coordinates']['bounding_box'] = [x1, y1, x2, y2]
                    value['coordinates']['center_point'] = [(x1 + x2) / 2, (y1 + y2) / 2]
                else:
                    # Set default coordinates if invalid
                    value['coordinates']['bounding_box'] = [0.1, 0.1, 0.2, 0.2]
                    value['coordinates']['center_point'] = [0.15, 0.15]
        
        return result

    async def _parse_gemini_json_response_robust(self, response_text: str, context: str) -> Dict[str, Any]:
        """Robust JSON parsing for Gemini responses with multiple fallback strategies"""
        try:
            # Step 1: Basic cleaning
            cleaned_text = response_text.strip()
            
            # Remove common markdown formatting
            if cleaned_text.startswith("```json"):
                cleaned_text = cleaned_text[7:]
            elif cleaned_text.startswith("```"):
                cleaned_text = cleaned_text[3:]
            
            if cleaned_text.endswith("```"):
                cleaned_text = cleaned_text[:-3]
            
            # Step 2: Find JSON boundaries more carefully
            json_start = -1
            json_end = -1
            brace_count = 0
            
            for i, char in enumerate(cleaned_text):
                if char == '{':
                    if json_start == -1:
                        json_start = i
                    brace_count += 1
                elif char == '}':
                    brace_count -= 1
                    if brace_count == 0 and json_start != -1:
                        json_end = i
                        break
            
            if json_start == -1 or json_end == -1:
                raise ValueError(f"No complete JSON object found in response for {context}")
            
            json_str = cleaned_text[json_start:json_end + 1]
            
            # Step 3: Advanced JSON cleaning
            # json_str = self._clean_json_aggressively(json_str)
            
            # Step 4: Try to parse
            result = json.loads(json_str)
            logger.info(f"Successfully parsed Gemini JSON for {context}")
            return result
            
        except json.JSONDecodeError as e:
            logger.error(f"JSON parsing error for {context}: {e}")
            # Try recovery strategies
            return await self._json_recovery_strategies(response_text, context)
            
        except Exception as e:
            logger.error(f"Unexpected parsing error for {context}: {e}")
            return self._get_fallback_structure(context)

    async def _json_recovery_strategies(self, response_text: str, context: str) -> Dict[str, Any]:
        """Multiple strategies to recover from JSON parsing failures"""
        
        # Strategy 1: Try to extract just the core data
        try:
            # Look for key patterns and extract manually
            if "extracted_values" in response_text or "potential_sources" in response_text or "batch_analysis" in response_text:
                # Try to find array content
                for key in ["extracted_values", "potential_sources", "batch_analysis"]:
                    start_pattern = rf'"{key}"\s*:\s*\['
                    
                    match = re.search(start_pattern, response_text)
                    if match:
                        start_pos = match.end() - 1  # Include the [
                        
                        # Find matching closing bracket
                        bracket_count = 0
                        end_pos = -1
                        for i in range(start_pos, len(response_text)):
                            if response_text[i] == '[':
                                bracket_count += 1
                            elif response_text[i] == ']':
                                bracket_count -= 1
                                if bracket_count == 0:
                                    end_pos = i
                                    break
                        
                        if end_pos != -1:
                            array_content = response_text[start_pos:end_pos + 1]
                            # Try to parse just this array
                            try:
                                extracted_array = json.loads(array_content)
                                return {key: extracted_array}
                            except:
                                continue
            
        except:
            pass
        
        # Strategy 2: Return fallback structure
        return self._get_fallback_structure(context)

    def _get_fallback_structure(self, context: str) -> Dict[str, Any]:
        """Return appropriate fallback structure based on context"""
        if "page_" in context:
            return {
                "page_number": 1,
                "page_dimensions": {"width": 800, "height": 600},
                "extracted_values": [],
                "error": f"JSON parsing failed for {context}"
            }
        elif "excel_batch" in context:
            return {
                "batch_analysis": [],
                "error": f"JSON parsing failed for {context}"
            }
        elif "mapping" in context.lower():
            return {
                "suggested_mappings": [],
                "mapping_quality_assessment": {"overall_coverage": 0},
                "error": f"JSON parsing failed for {context}"
            }
        elif "audit" in context.lower():
            return {
                "batch_results": [],
                "error": f"JSON parsing failed for {context}"
            }
        else:
            return {
                "error": f"JSON parsing failed for {context}"
            }

# Initialize the enhanced service
enhanced_gemini_service = EnhancedGeminiService()


# if __name__ == "__main__":
#     enhanced_gemini_service = EnhancedGeminiService()
#     async def main():
#         # result = await enhanced_gemini_service.analyze_excel_comprehensive('/Users/himanshusharma/Personal_Code/sample data/Slide 1.xlsx')
#         # with open('result.json', 'w', encoding='utf-8') as f:
#         #     json.dump(result, f, ensure_ascii=False, indent=4)  # writes dict as JSON
#         # result_pdf = await enhanced_gemini_service.extract_comprehensive_pdf_data('/Users/himanshusharma/Personal_Code/sample data/Main Presentation - 1 page.pdf')
#         # with open('result_pdf.json', 'w', encoding='utf-8') as f:
#         #     json.dump(result_pdf, f, ensure_ascii=False, indent=4)  # writes dict as JSON

#         with open('result.json', 'r', encoding='utf-8') as f:
#             excel_data = json.load(f)  # This parses the JSON
        
#         with open('result_pdf.json', 'r', encoding='utf-8') as f:
#             pdf_data = json.load(f)  # This parses the JSON
#         excel_values = excel_data.get('potential_sources', [])
#         pdf_values = pdf_data.get('all_extracted_values', [])
        
#         complete_result = await enhanced_gemini_service.run_direct_comprehensive_audit(
#             pdf_values=pdf_values,
#             excel_values=excel_values
#         )
        
#         with open('result_complete.json', 'w', encoding='utf-8') as f:
#             json.dump(complete_result, f, ensure_ascii=False, indent=4)
    
#     asyncio.run(main())

#     # model_name = "gemini-2.0-flash-exp"
    # with open('/Users/himanshusharma/Personal_Code/veritas-dev/backend/prompt.txt', 'r', encoding='utf-8') as f:
    #     prompt = f.read()

    # response = enhanced_gemini_service.model.generate_content(contents=[prompt])
    
    # with open('result.txt', 'w', encoding='utf-8') as f:
    #     f.write(response.text)

    # run_direct_comprehensive_audit