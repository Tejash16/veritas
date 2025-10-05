import google.generativeai as genai
from typing import Dict, Any, List, Tuple, Optional
import json
import base64
import structlog
from PIL import Image
import io
import time
import faiss
import asyncio
import os
import fitz  # PyMuPDF
import pandas as pd
from decouple import config
import re
from datetime import datetime
import math
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import hashlib
from openpyxl import load_workbook
from app.services.audit.pdf_analysis import PdfAnalysisService
from app.services.audit.excel_analysis import ExcelAuditSystem
from app.services.audit.fuzzy_matching import FuzzyMatchService

# Configure logging
logger = structlog.get_logger()

class EnhancedGeminiService:

    def __init__(self):
        api_key = config('GOOGLE_API_KEY', default=None)
        if not api_key:
            raise ValueError("GOOGLE_API_KEY is required for Gemini 2.5 Pro")
        
        genai.configure(api_key=api_key)
        # Using more capable models
        self.model = genai.GenerativeModel('gemini-2.0-flash-exp')
        self.ai_enabled = True
        self.vector_db_path = None
        # Improved embedding model
        self.embedding_model_name = 'models/embedding-001'
        self.context_json_path = "faiss_db/contexts.json"
        self.fuzzy_match_service = FuzzyMatchService()
        
        logger.info("Enhanced Gemini Service initialized with improved extraction algorithms")

    async def extract_comprehensive_pdf_data(self, pdf_path: str) -> Dict[str, Any]:
        """Enhanced PDF extraction with multi-modal approach"""
        logger.info(f"Starting enhanced PDF extraction: {pdf_path}")
        
        try:
            pd = PdfAnalysisService()
            result = await pd.extract_comprehensive_pdf_data(pdf_path)
            
            return result
            
        except Exception as e:
            logger.error(f"PDF extraction failed: {e}")
            raise

    async def analyze_excel_comprehensive(self, excel_path: str) -> Dict[str, Any]:
        """Enhanced Excel analysis with better structure detection"""
        logger.info(f"Starting enhanced Excel analysis: {excel_path}")
        try:
            ea = ExcelAuditSystem()
            result = ea.analyse_excel_comprehensive(excel_path)
            # print(result)
            self.vector_db_path = result.faiss_index_path
            analysed_values = result.analysedValues
            logger.info(f"Excel analysis completed: {analysed_values} potential sources identified")
            result = {
                "workbook_summary": {
                    "total_sheets_processed": '',
                    "total_potential_sources": '',
                    "analysis_timestamp": datetime.utcnow().isoformat(),
                },
                "potential_sources": analysed_values,
                "sheet_analyses": ''
            }
            return result
        except Exception as e:
            logger.error(f"Excel extraction failed: {e}")
            raise

    async def run_direct_comprehensive_audit(self, pdf_json_data: dict, batch_size: int = 10):
        """Process PDF values against FAISS vector database and send to LLM for final validation"""
        pdf_values = pdf_json_data
        faiss_index_path = 'faiss_db/faiss_index.index'
        logger.info(f"Processing {len(pdf_values)} PDF values against vector database")

        # Load FAISS index and contexts
        if not self.load_vector_database(faiss_index_path, context_json_path=self.context_json_path):
            return {"error": "Failed to load vector database"}

        all_results = [] 
        total_batches = math.ceil(len(pdf_values) / batch_size)

        for batch_num in range(total_batches):
            start_idx = batch_num * batch_size
            end_idx = min((batch_num + 1) * batch_size, len(pdf_values))
            batch = pdf_values[start_idx:end_idx]

            logger.info(f"Processing batch {batch_num + 1}/{total_batches} ({len(batch)} values)")
            batch_results = await self._process_vector_audit_batch(batch, batch_num + 1, total_batches)
            all_results.extend(batch_results)

            await asyncio.sleep(1)  # rate limiting

        return {
            "summary": self._calculate_vector_audit_summary(all_results),
            "detailed_results": all_results,
            "total_processed": len(pdf_values)
        }

    def load_vector_database(self, faiss_index_path: str, context_json_path: str) -> bool:
        """Load FAISS index and contexts from disk"""
        try:
            self.faiss_index = faiss.read_index(faiss_index_path)
            logger.info(f"Loaded FAISS index with {self.faiss_index.ntotal} vectors")

            with open(context_json_path, "r", encoding="utf-8") as f:
                self.contexts = json.load(f)
            logger.info(f"Loaded {len(self.contexts)} contexts from {context_json_path}")
            # Pre-load Excel data for fuzzy matching
            if not self.fuzzy_match_service.load_excel_data(self.contexts):
                logger.warning("Fuzzy matching data loading failed, but continuing without fuzzy matching")
            
            return True
        except Exception as e:
            logger.error(f"Failed to load vector database: {e}")
            return False

    async def _process_vector_audit_batch(self, pdf_batch: list, batch_num: int, total_batches: int):
        """Process a batch of PDF values using fuzzy + vector search and LLM validation"""
        # Perform fuzzy matching first
        fuzzy_results = await self._perform_fuzzy_matching(pdf_batch)
        
        enhanced_batch = []
        
        for i, pdf_value in enumerate(pdf_batch):
            # Get fuzzy matches for this value
            fuzzy_matches = []
            if i < len(fuzzy_results) and fuzzy_results[i]:
                fuzzy_matches = fuzzy_results[i].get('fuzzy_matches', [])
            
            # Build query text for vector search (existing code)
            business_context = pdf_value.get('business_context', {})
            query_text = f"Value: {pdf_value['value']} | Context: {business_context.get('semantic_meaning', '')} | Category: {business_context.get('business_category', '')} | Type: {business_context.get('calculation_type', '')}"

            # Generate embedding and search FAISS (existing code)
            embedding = await self._generate_embedding(query_text)
            D, I = self.faiss_index.search(np.array([embedding]).astype('float32'), k=5)
            
            vector_matches = []
            for idx, score in zip(I[0], D[0]):
                if idx < len(self.contexts):
                    ctx = self.contexts[idx]
                    vector_matches.append({
                        "excel_value": ctx.get("value"),
                        "excel_location": ctx.get("cell_address"),
                        "excel_context": ctx.get("full_context"),
                        "confidence": float(score),
                        "match_type": "vector_match",
                        "business_context": ctx.get("table_title")
                    })

            enhanced_batch.append({
                "id": pdf_value.get('id'),
                "value": pdf_value.get('value'),
                "context": business_context.get('semantic_meaning', ''),
                "business_category": business_context.get('business_category', ''),
                "data_type": pdf_value.get('data_type'),
                "fuzzy_matches": fuzzy_matches,  # Add fuzzy matches
                "vector_matches": vector_matches  # Existing vector matches
            })

        # Update the prompt to include both fuzzy and vector matches
        prompt = f"""
    You are auditing presentation values against Excel source data using FUZZY matches and VECTOR database matches.

    BATCH {batch_num}/{total_batches} PDF VALUES TO VALIDATE:
    {json.dumps(enhanced_batch, indent=1)}

    For EACH PDF value (process them in order):
    1. FIRST check FUZZY matches (exact, numeric, similar values)
    2. THEN check VECTOR matches (semantic/contextual similarity)  
    3. If no direct match, check if it's a derived metric (growth rate, percentage, ratio)
    4. For derived metrics, attempt calculation using relevant Excel data
    5. Assign appropriate validation status

    PRIORITIZATION:
    - Use fuzzy matches first when confidence > 0.9
    - Fall back to vector matches when fuzzy matches are weak
    - Consider both match types in your reasoning

    Return ONLY valid JSON: 
    {{
        "batch_results": [
            {{
                "pdf_value_id": "value_1_001",
                "pdf_value": "2759",
                "pdf_context": "FY25 revenue growth",
                "validation_status": "matched|mismatched|unverifiable",
                "excel_match": {{
                    "source_cell": "Sheet!B5",
                    "Excel_Cell_used_for_match": "B5", (if source cell is more then one cell follow B5!O10 pattern)
                    "Source_Sheet": "Sheet",
                    "excel_value": "2759",
                    "match_confidence": 0.95,
                    "calculation_basis": "direct_match|calculated|inferred",
                    "match_source": "fuzzy|vector|both"
                }},
                "confidence": 0.95,
                "audit_reasoning": "Detailed explanation including which match type was used and calculation steps if applicable"
            }}
        ]
    }}


    When comparing values:
    - Round Excel floats to nearest 2 decimal places.
    - Convert decimals to percentages if context mentions "%".
    - Treat values within ±5% difference as equivalent matches.
    - Use Excel context (row/column headers, table titles) to determine if a PDF label corresponds to that Excel value.

    Validation Status Guide(Only these 3 validation status is allowed):
    - matched: Exact or equivalent value found, match_confidence greater than 0.90. Calculated values also have validation status as matched. PDF or Excel match also fall under this category. excel_match or pdf_match should also have validation status as matched
    - mismatched: Different values but related context  
    - unverifiable: Cannot verify relationship. Also when the value is only present in pdf

    IMPORTANT RULES:
    1. Return EXACTLY {len(pdf_batch)} results, one for each PDF value IN ORDER
    2. For calculated values, set the validation_status as matched and show the calculation in audit_reasoning
    3. Return ONLY valid JSON, no other text
"""
        
        # with open('prompt.txt','w',encoding='utf-8') as f:
        #     f.write(prompt)
        #     f.write("====================================")

        try:
            response = self.model.generate_content(prompt)
            result = await self._parse_gemini_json_response_robust(response.text, f"vector_audit_batch_{batch_num}")
            # with open('result.txt','w',encoding='utf-8') as f:
            #     f.write(str(result))
            #     f.write("====================================")
            
            batch_results = result.get("batch_results", [])

            # Ensure correct number of results
            final_results = []
            for i, pdf_value in enumerate(pdf_batch):
                if i < len(batch_results):
                    final_results.append(batch_results[i])
                else:
                    # fallback if Gemini response is incomplete
                    final_results.append({
                        "pdf_value_id": pdf_value.get('id'),
                        "pdf_value": pdf_value.get('value'),
                        "validation_status": "unverifiable",
                        "confidence": 0.0
                    })

            logger.info(f"Vector batch {batch_num}: Processed {len(final_results)} PDF values")
            return final_results

        except Exception as e:
            logger.error(f"Vector audit batch {batch_num} failed: {e}")
            return [{
                "pdf_value_id": pdf_value.get('id'),
                "pdf_value": pdf_value.get('value'),
                "validation_status": "unverifiable",
                "confidence": 0.0,
                "error": str(e)
            } for pdf_value in pdf_batch]

    async def _perform_fuzzy_matching(self, pdf_batch: List[Dict]) -> List[Dict]:
            """Perform fuzzy matching before vector matching"""
            try:
                logger.info(f"Starting fuzzy matching for {len(pdf_batch)} values")
                
                # Perform fuzzy matching using the pre-initialized service
                fuzzy_results = await self.fuzzy_match_service.fuzzy_match_batch(pdf_batch)
                
                # Log summary
                summary = self.fuzzy_match_service.get_matching_summary(fuzzy_results)
                logger.info(f"Fuzzy matching completed: {summary}")
                
                return fuzzy_results
                
            except Exception as e:
                logger.error(f"Fuzzy matching failed: {e}")
                return []

    async def _generate_embedding(self, text: str):
        """Generate embedding using Gemini API"""
        result = genai.embed_content(model="models/embedding-001", content=text)
        return np.array(result['embedding']).astype('float32')

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

    def _calculate_vector_audit_summary(self, results: list) -> dict:
        """Simple summary statistics"""
        total = len(results)
        matched = sum(1 for r in results if r.get("validation_status") == "matched")
        mismatched = sum(1 for r in results if r.get("validation_status") == "mismatched")
        unverifiable = total - matched - mismatched
        return {
            "total": total,
            "matched": matched,
            "mismatched": mismatched,
            "unverifiable": unverifiable
        }

enhanced_gemini_service = EnhancedGeminiService()

# if __name__ == "__main__":
#     enhanced_gemini_service = EnhancedGeminiService()
#     async def main():
#         # result_pdf = await enhanced_gemini_service.extract_comprehensive_pdf_data('/Users/himanshusharma/Downloads/ABHI IR Q1FY26 v12.pdf')
#         # with open('result.json', 'w', encoding='utf-8') as f:
#         #     json.dump(result_pdf, f, ensure_ascii=False, indent=4)
        
#         # result = await enhanced_gemini_service.analyze_excel_comprehensive('/Users/himanshusharma/Downloads/investor deck working_Q1 FY26.xlsx')
#         # print(result)
#         # result = await enhanced_gemini_service.analyze_excel_comprehensive('/Users/himanshusharma/Personal_Code/sample data/Slide 1.xlsx')
#         # print(result)

#         with open('/Users/himanshusharma/Personal_Code/veritas-dev/result.json', 'r', encoding='utf-8') as f:
#             pdf_data = json.load(f)
#         result = await enhanced_gemini_service.run_direct_comprehensive_audit(pdf_data,faiss_index_path='/Users/himanshusharma/Personal_Code/veritas-dev/faiss_db/faiss_index.index')
#         with open('result_complete.json', 'w', encoding='utf-8') as f:
#             json.dump(result, f, ensure_ascii=False, indent=4)
#     asyncio.run(main())
