from pydantic_settings import BaseSettings
from typing import Optional
import os

class Settings(BaseSettings):
    # Database
    DATABASE_URL: str = "sqlite:///./veritas_enhanced.db"
    
    # Security
    SECRET_KEY: str = "your-secret-key-change-in-production"
    ALGORITHM: str = "HS256"
    ACCESS_TOKEN_EXPIRE_MINUTES: int = 30
    
    # AI Services
    GOOGLE_API_KEY: Optional[str] = None
    OPENAI_API_KEY: Optional[str] = None
    
    # COMPREHENSIVE AI Settings - Updated for full extraction
    enhancement_level: str = "comprehensive"  # Changed from "full" to "comprehensive"
    ai_processing_timeout: int = 600  # Increased from 300 to 600 seconds (10 minutes)
    semantic_similarity_threshold: float = 0.6
    
    # COMPREHENSIVE Excel Processing Settings
    max_sheets_per_workbook: int = 50      # Up from previous limits
    max_rows_per_sheet: int = 2000         # Up from 50
    max_cols_per_sheet: int = 200          # Up from 20
    max_cells_per_batch: int = 200         # Intelligent batching
    enable_comprehensive_extraction: bool = True
    
    # COMPREHENSIVE PDF Processing Settings  
    max_pdf_pages: int = 100               # Support larger presentations
    pdf_image_quality: float = 2.0         # High quality for better AI extraction
    enable_coordinate_mapping: bool = True
    
    # AI Batching and Rate Limiting
    gemini_batch_size: int = 200           # Cells per Gemini API call
    gemini_rate_limit_delay: float = 1.0   # Seconds between API calls
    max_concurrent_ai_requests: int = 3    # Parallel processing limit
    
    # Memory Management
    enable_memory_efficient_processing: bool = True
    max_memory_usage_mb: int = 1024        # 1GB memory limit
    
    # File Storage - Updated for larger files
    UPLOAD_DIR: str = "./uploads"
    MAX_FILE_SIZE: int = 200 * 1024 * 1024  # Increased to 200MB from 100MB
    MAX_EXCEL_FILE_SIZE: int = 150 * 1024 * 1024  # 150MB for Excel files
    MAX_PDF_FILE_SIZE: int = 100 * 1024 * 1024    # 100MB for PDF files
    
    # Redis
    REDIS_URL: str = "redis://localhost:6379"
    
    # Metrics
    ENABLE_METRICS: bool = True
    METRICS_PORT: int = 8001
    
    # Logging - Enhanced for comprehensive processing
    LOG_LEVEL: str = "INFO"
    ENABLE_DETAILED_LOGGING: bool = True
    LOG_AI_RESPONSES: bool = False  # Set to True for debugging AI responses
    
    # Performance Settings
    ENABLE_CACHING: bool = True
    CACHE_TTL_SECONDS: int = 3600
    
    # API Response Limits - Updated for comprehensive data
    MAX_API_RESPONSE_SIZE_MB: int = 50     # Increased from typical limits
    MAX_EXCEL_VALUES_IN_RESPONSE: int = 10000  # No artificial limits
    MAX_PDF_VALUES_IN_RESPONSE: int = 1000     # Support large presentations
    
    # Validation Settings
    VALIDATION_CONFIDENCE_THRESHOLD: float = 0.3  # Lower threshold for comprehensive coverage
    ENABLE_FUZZY_MATCHING: bool = True
    FUZZY_MATCH_THRESHOLD: float = 0.8
    
    # Direct Validation Settings
    ENABLE_DIRECT_VALIDATION: bool = True
    DIRECT_VALIDATION_BATCH_SIZE: int = 5  # PDF values per audit batch
    MAX_EXCEL_VALUES_FOR_COMPARISON: int = 100  # Excel values to compare against per batch
    
    # Error Handling
    MAX_RETRIES: int = 3
    RETRY_DELAY_SECONDS: float = 2.0
    CONTINUE_ON_SHEET_ERRORS: bool = True  # Don't fail entire workbook if one sheet fails
    
    # Development and Testing
    MOCK_AI_RESPONSES: bool = False  # For testing without AI API
    ENABLE_DEBUG_ENDPOINTS: bool = False
    
    class Config:
        env_file = ".env"
        extra = "ignore"  # Ignore extra fields

# Settings validation and computed properties
class ComprehensiveSettings(Settings):
    """Extended settings with computed properties for comprehensive processing"""
    
    @property
    def is_comprehensive_mode(self) -> bool:
        """Check if comprehensive extraction is enabled"""
        return self.enable_comprehensive_extraction and self.enhancement_level == "comprehensive"
    
    @property
    def effective_ai_timeout(self) -> int:
        """Get effective AI timeout based on comprehensive mode"""
        if self.is_comprehensive_mode:
            return max(self.ai_processing_timeout, 600)  # Minimum 10 minutes for comprehensive
        return self.ai_processing_timeout
    
    @property
    def max_total_cells_to_process(self) -> int:
        """Calculate maximum total cells that can be processed"""
        return self.max_sheets_per_workbook * self.max_rows_per_sheet * self.max_cols_per_sheet
    
    @property
    def gemini_api_settings(self) -> dict:
        """Get optimized Gemini API settings for comprehensive processing"""
        return {
            "batch_size": self.gemini_batch_size,
            "rate_limit_delay": self.gemini_rate_limit_delay,
            "max_concurrent_requests": self.max_concurrent_ai_requests,
            "timeout": self.effective_ai_timeout
        }
    
    def get_processing_limits_summary(self) -> dict:
        """Get a summary of all processing limits for logging"""
        return {
            "comprehensive_mode": self.is_comprehensive_mode,
            "max_sheets": self.max_sheets_per_workbook,
            "max_rows_per_sheet": self.max_rows_per_sheet,
            "max_cols_per_sheet": self.max_cols_per_sheet,
            "max_file_size_mb": self.MAX_FILE_SIZE // (1024 * 1024),
            "max_excel_values": self.MAX_EXCEL_VALUES_IN_RESPONSE,
            "max_pdf_values": self.MAX_PDF_VALUES_IN_RESPONSE,
            "ai_timeout_seconds": self.effective_ai_timeout
        }

settings = ComprehensiveSettings()

# Log the comprehensive settings on startup
def log_comprehensive_settings():
    """Log comprehensive processing settings for verification"""
    import structlog
    logger = structlog.get_logger()
    
    if settings.is_comprehensive_mode:
        logger.info("🚀 COMPREHENSIVE EXTRACTION MODE ENABLED")
        logger.info("Processing limits summary:", **settings.get_processing_limits_summary())
        logger.info("All artificial limits have been removed for maximum data extraction")
    else:
        logger.warning("⚠️ Comprehensive extraction not enabled - check settings")

# Environment variable validation
def validate_comprehensive_settings():
    """Validate that all required settings are properly configured"""
    import structlog
    logger = structlog.get_logger()
    
    issues = []
    
    # Check AI API key
    if not settings.GOOGLE_API_KEY:
        issues.append("GOOGLE_API_KEY is not set - AI extraction will fail")
    
    # Check if comprehensive settings are reasonable
    if settings.max_sheets_per_workbook < 10:
        issues.append(f"max_sheets_per_workbook ({settings.max_sheets_per_workbook}) seems low for comprehensive extraction")
    
    if settings.max_rows_per_sheet < 1000:
        issues.append(f"max_rows_per_sheet ({settings.max_rows_per_sheet}) seems low for comprehensive extraction")
    
    if settings.ai_processing_timeout < 300:
        issues.append(f"ai_processing_timeout ({settings.ai_processing_timeout}s) might be too low for comprehensive processing")
    
    # Log issues
    if issues:
        logger.warning("⚠️ Configuration issues detected:")
        for issue in issues:
            logger.warning(f"  - {issue}")
    else:
        logger.info("✅ Comprehensive extraction configuration validated successfully")
    
    return len(issues) == 0

# Initialize on import
if __name__ != "__main__":
    log_comprehensive_settings()
    validate_comprehensive_settings()