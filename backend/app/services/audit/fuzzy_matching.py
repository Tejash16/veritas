# fuzzy_match_service.py
import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz, process
import structlog
from typing import Dict, Any, List, Tuple, Optional
import json
from datetime import datetime

logger = structlog.get_logger()

class FuzzyMatchService:
    def __init__(self, similarity_threshold: float = 0.85):
        self.similarity_threshold = similarity_threshold
        self.excel_data = None
        logger.info("Fuzzy Match Service initialized")
    
    def load_excel_data(self, contexts: List[Dict]) -> bool:
        """Load Excel data from vector DB contexts for fuzzy matching"""
        try:
            self.excel_data = []
            for ctx in contexts:
                self.excel_data.append({
                    'value': ctx.get('value', ''),
                    'cell_address': ctx.get('cell_address', ''),
                    'full_context': ctx.get('full_context', ''),
                    'table_title': ctx.get('table_title', ''),
                    'sheet_name': ctx.get('sheet_name', ''),
                    'data_type': ctx.get('data_type', '')
                })
            logger.info(f"Loaded {len(self.excel_data)} Excel records for fuzzy matching")
            return True
        except Exception as e:
            logger.error(f"Failed to load Excel data for fuzzy matching: {e}")
            return False
    
    async def fuzzy_match_batch(self, pdf_batch: List[Dict]) -> List[Dict]:
        """Perform fuzzy matching on a batch of PDF values"""
        if not self.excel_data:
            logger.warning("No Excel data loaded for fuzzy matching")
            return []
        
        fuzzy_results = []
        
        for pdf_value in pdf_batch:
            pdf_val_str = str(pdf_value.get('value', ''))
            pdf_context = pdf_value.get('business_context', {}).get('semantic_meaning', '')
            pdf_category = pdf_value.get('business_context', {}).get('business_category', '')
            
            matches = await self._find_fuzzy_matches(
                pdf_val_str, 
                pdf_context, 
                pdf_category
            )
            
            fuzzy_results.append({
                "id": pdf_value.get('id'),
                "value": pdf_value.get('value'),
                "context": pdf_context,
                "business_category": pdf_category,
                "data_type": pdf_value.get('data_type'),
                "fuzzy_matches": matches
            })
        
        return fuzzy_results
    
    async def _find_fuzzy_matches(self, pdf_value: str, pdf_context: str, pdf_category: str) -> List[Dict]:
        """Find fuzzy matches for a single PDF value"""
        matches = []
        
        try:
            # Direct value matching
            for excel_item in self.excel_data:
                excel_val = str(excel_item.get('value', ''))
                
                # Skip empty values
                if not excel_val or not pdf_value:
                    continue
                
                # Try exact match first
                if self._is_exact_match(pdf_value, excel_val):
                    matches.append({
                        "excel_value": excel_val,
                        "excel_location": excel_item.get('cell_address'),
                        "excel_context": excel_item.get('full_context'),
                        "confidence": 1.0,
                        "match_type": "exact_value",
                        "business_context": excel_item.get('table_title'),
                        "matching_algorithm": "exact"
                    })
                    continue
                
                # Try numeric matching for numerical values
                if self._is_numeric_match(pdf_value, excel_val):
                    matches.append({
                        "excel_value": excel_val,
                        "excel_location": excel_item.get('cell_address'),
                        "excel_context": excel_item.get('full_context'),
                        "confidence": 0.95,
                        "match_type": "numeric",
                        "business_context": excel_item.get('table_title'),
                        "matching_algorithm": "numeric"
                    })
                    continue
                
                # Try fuzzy string matching
                similarity = fuzz.ratio(str(pdf_value).lower(), str(excel_val).lower()) / 100.0
                
                if similarity >= self.similarity_threshold:
                    matches.append({
                        "excel_value": excel_val,
                        "excel_location": excel_item.get('cell_address'),
                        "excel_context": excel_item.get('full_context'),
                        "confidence": similarity,
                        "match_type": "fuzzy_value",
                        "business_context": excel_item.get('table_title'),
                        "matching_algorithm": "fuzzy_string"
                    })
            
            # Context-based matching if no direct matches found
            if not matches and pdf_context:
                matches.extend(await self._context_based_matching(pdf_value, pdf_context, pdf_category))
            
            # Sort by confidence and return top 5 matches
            matches.sort(key=lambda x: x['confidence'], reverse=True)
            return matches[:5]
            
        except Exception as e:
            logger.error(f"Fuzzy matching error for value '{pdf_value}': {e}")
            return []
    
    def _is_exact_match(self, pdf_value: str, excel_value: str) -> bool:
        """Check for exact match with normalization"""
        try:
            pdf_clean = str(pdf_value).strip().lower()
            excel_clean = str(excel_value).strip().lower()
            
            # Remove common formatting differences
            pdf_clean = pdf_clean.replace(',', '').replace(' ', '')
            excel_clean = excel_clean.replace(',', '').replace(' ', '')
            
            return pdf_clean == excel_clean
        except:
            return False
    
    def _is_numeric_match(self, pdf_value: str, excel_value: str) -> bool:
        """Check for numeric equivalence with tolerance"""
        try:
            # Try to convert to numbers
            pdf_num = float(str(pdf_value).replace(',', '').replace('%', '').strip())
            excel_num = float(str(excel_value).replace(',', '').replace('%', '').strip())
            
            # Check if they're within 1% of each other
            if pdf_num == 0 and excel_num == 0:
                return True
            elif pdf_num == 0 or excel_num == 0:
                return False
            
            difference = abs(pdf_num - excel_num) / max(abs(pdf_num), abs(excel_num))
            return difference <= 0.01  # 1% tolerance
            
        except (ValueError, TypeError):
            return False
    
    async def _context_based_matching(self, pdf_value: str, pdf_context: str, pdf_category: str) -> List[Dict]:
        """Perform context-based matching when direct value matching fails"""
        context_matches = []
        
        try:
            for excel_item in self.excel_data:
                excel_context = excel_item.get('full_context', '') + ' ' + excel_item.get('table_title', '')
                
                # Calculate context similarity
                context_similarity = fuzz.partial_ratio(
                    str(pdf_context).lower(), 
                    str(excel_context).lower()
                ) / 100.0
                
                if context_similarity >= 0.7:  # Context similarity threshold
                    context_matches.append({
                        "excel_value": excel_item.get('value'),
                        "excel_location": excel_item.get('cell_address'),
                        "excel_context": excel_item.get('full_context'),
                        "confidence": context_similarity * 0.8,  # Weight context matches lower
                        "match_type": "context_based",
                        "business_context": excel_item.get('table_title'),
                        "matching_algorithm": "context_similarity"
                    })
            
            return context_matches
            
        except Exception as e:
            logger.error(f"Context-based matching error: {e}")
            return []
    
    def get_matching_summary(self, fuzzy_results: List[Dict]) -> Dict[str, Any]:
        """Generate summary statistics for fuzzy matching"""
        total_values = len(fuzzy_results)
        values_with_matches = sum(1 for item in fuzzy_results if item.get('fuzzy_matches'))
        total_matches = sum(len(item.get('fuzzy_matches', [])) for item in fuzzy_results)
        
        return {
            "total_pdf_values": total_values,
            "values_with_fuzzy_matches": values_with_matches,
            "total_fuzzy_matches": total_matches,
            "match_rate": values_with_matches / total_values if total_values > 0 else 0
        }