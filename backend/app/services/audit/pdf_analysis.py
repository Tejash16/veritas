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

class PdfAnalysisService:
    def __init__(self):
        # Get API key from environment
        api_key = config('GOOGLE_API_KEY', default=None)
        
        if not api_key:
            raise ValueError("GOOGLE_API_KEY is required for Gemini 2.5 Pro")
        
        genai.configure(api_key=api_key)
        # self.model = genai.GenerativeModel('gemini-2.0-flash-exp')
        self.model = genai.GenerativeModel('gemini-2.5-flash-lite')
        self.ai_enabled = True

        logger.info("Enhanced Gemini 2.5 Pro PDF Service initialized with comprehensive extraction settings")


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
                if len(page_data['extracted_values']) == 0:
                    continue
                page_analyses.append(page_data)
                
                # Rate limiting for Gemini API
                # await asyncio.sleep(1)
            
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
            Analyze this presentation slide (page {page_num}) and extract ONLY meaningful business-related numerical data. 
            Focus strictly on quantitative or financial metrics that convey insight — such as revenues, profits, growth %, market share, and performance indicators.

            ❌ HARD DO NOT EXTRACT:
            - Section numbers, slide numbers, or page numbers (e.g., "1", "2", "Page 3")
            - Bullet or list numbers
            - Decorative numbers or formatting numbers (asterisks, superscripts, footnotes)
            - Standalone fiscal years or quarters (e.g., "FY26", "Q1 FY26") unless directly tied to a numeric metric (e.g., "Revenue grew 40% in FY26")
            - Any number that exists purely as a reference, label, or metadata

            ✅ ONLY EXTRACT:
            - Financial, market, or operational metrics with clear business meaning
            - Percentages, ratios, and growth figures
            - Quantitative data directly tied to tables, charts, or performance indicators
            - Dates or periods only if they are attached to a numeric metric (e.g., "Revenue in Q1 FY26: $10M")

            For each number, provide:
            - Exact value as displayed
            - Business context
            - Normalized coordinates [x1, y1, x2, y2] on 0–1 scale
            - Data type classification
            - Only include data points that provide measurable business insight

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
            If the only number on a page appears isolated (e.g., "1" or "2" in a corner or footer) 
            with context related to page number or slide number, ignore that
            and return a blank json {{}}
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

        synthesis = {
                "document_summary": {
                    "total_pages": len(page_analyses),
                    "document_type": "financial_presentation"
                },
                "all_extracted_values": all_values,
                "extraction_quality_metrics": {
                    "total_values_extracted": len(all_values),
                    "overall_confidence": 0.8
                }
            }
        logger.info("Document synthesis completed successfully")
        return synthesis



if __name__ == "__main__":
    async def main():
        pd = PdfAnalysisService()
        result = await pd.extract_comprehensive_pdf_data('/Users/himanshusharma/Downloads/ABHI IR Q1FY26 v12.pdf')
        # result = await pd.extract_comprehensive_pdf_data('/Users/himanshusharma/Personal_Code/sample data/Main Presentation - 1 page.pdf')
        with open('result.json', 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=4)
    asyncio.run(main())