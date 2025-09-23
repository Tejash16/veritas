// src/components/Upload/EnhancedMappingValidation.jsx
import React, { useState, useEffect } from 'react';
import { useParams, useNavigate, useLocation } from 'react-router-dom';
import { Check, X, Edit3, Eye, AlertTriangle, TrendingUp, ArrowRight } from 'lucide-react';
import toast from 'react-hot-toast';

import { apiService } from '../../services/api';
import EnhancedDocumentPreview from '../Preview/EnhancedDocumentPreview';
import LoadingSpinner from '../Common/LoadingSpinner';

const EnhancedMappingValidation = () => {
  const { sessionId } = useParams();
  const navigate = useNavigate();
  const location = useLocation();
  
  const [documentData, setDocumentData] = useState(null);
  const [extractedValues, setExtractedValues] = useState([]);
  const [mappingSuggestions, setMappingSuggestions] = useState([]);
  const [selectedValueId, setSelectedValueId] = useState(null);
  const [loading, setLoading] = useState(true);
  const [validationMode, setValidationMode] = useState('review'); // 'review' | 'validate' | 'audit'

  useEffect(() => {
    loadEnhancedData();
  }, [sessionId]);

  const loadEnhancedData = async () => {
    try {
      // Load comprehensive extraction data
      const response = await apiService.getEnhancedExtractionData(sessionId);
      
      setDocumentData(response.document_data);
      setExtractedValues(response.extracted_values);
      setMappingSuggestions(response.mapping_suggestions);
      
    } catch (error) {
      console.error('Failed to load enhanced data:', error);
      toast.error('Failed to load extraction data');
    } finally {
      setLoading(false);
    }
  };

  const handleValueEdit = async (valueId, newData) => {
    try {
      await apiService.updateExtractedValue(sessionId, valueId, newData);
      
      setExtractedValues(prev => 
        prev.map(val => 
          val.id === valueId 
            ? { ...val, ...newData, user_modified: true }
            : val
        )
      );
      
      toast.success('Value updated successfully');
    } catch (error) {
      toast.error('Failed to update value');
    }
  };

  const handleMappingConfirm = async (mappingId) => {
    try {
      await apiService.confirmMapping(sessionId, mappingId);
      
      setMappingSuggestions(prev =>
        prev.map(mapping =>
          mapping.mapping_id === mappingId
            ? { ...mapping, user_confirmed: true, confidence: Math.max(mapping.confidence, 0.9) }
            : mapping
        )
      );
      
      toast.success('Mapping confirmed');
    } catch (error) {
      toast.error('Failed to confirm mapping');
    }
  };

  const handleMappingReject = (mappingId) => {
    setMappingSuggestions(prev =>
      prev.filter(mapping => mapping.mapping_id !== mappingId)
    );
    toast.info('Mapping removed');
  };

  const startAdvancedAudit = async () => {
    const confirmedMappings = mappingSuggestions.filter(m => m.user_confirmed);
    
    if (confirmedMappings.length === 0) {
      toast.error('Please confirm at least one mapping before starting audit');
      return;
    }

    setLoading(true);

    try {
      const auditRequest = {
        session_id: sessionId,
        confirmed_mappings: confirmedMappings,
        extracted_values: extractedValues,
        validation_mode: 'comprehensive'
      };

      const auditResponse = await apiService.runAdvancedAudit(auditRequest);
      
      toast.success('Advanced audit completed!');
      navigate(`/audit/${sessionId}`, {
        state: {
          auditResults: auditResponse.audit_results,
          extractionData: { documentData, extractedValues, mappingSuggestions }
        }
      });

    } catch (error) {
      console.error('Advanced audit failed:', error);
      toast.error('Audit failed: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-96">
        <LoadingSpinner size="xl" text="Loading enhanced validation interface..." />
      </div>
    );
  }

  return (
    <div className="max-w-7xl mx-auto space-y-6">
      {/* Header */}
      <div className="bg-white shadow rounded-lg p-6">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold text-gray-900">Enhanced Validation</h1>
            <p className="text-gray-600 mt-1">
              Interactive validation with AI-powered mapping suggestions
            </p>
          </div>
          
          <div className="flex items-center space-x-4">
            <div className="flex bg-gray-100 rounded-lg p-1">
              {['review', 'validate', 'audit'].map((mode) => (
                <button
                  key={mode}
                  onClick={() => setValidationMode(mode)}
                  className={`px-3 py-1 text-sm font-medium rounded-md capitalize ${
                    validationMode === mode
                      ? 'bg-white text-blue-600 shadow-sm'
                      : 'text-gray-600 hover:text-gray-900'
                  }`}
                >
                  {mode}
                </button>
              ))}
            </div>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Document Preview */}
        <div className="space-y-4">
          <EnhancedDocumentPreview
            documentData={documentData}
            extractedValues={extractedValues}
            onValueEdit={handleValueEdit}
            onValueSelect={setSelectedValueId}
            selectedValueId={selectedValueId}
          />
          
          {/* Extraction Summary */}
          <ExtractionSummary extractedValues={extractedValues} />
        </div>

        {/* Validation Panel */}
        <div className="space-y-4">
          {validationMode === 'review' && (
            <ValueReviewPanel
              extractedValues={extractedValues}
              selectedValueId={selectedValueId}
              onValueSelect={setSelectedValueId}
              onValueEdit={handleValueEdit}
            />
          )}
          
          {validationMode === 'validate' && (
            <MappingValidationPanel
              mappingSuggestions={mappingSuggestions}
              onMappingConfirm={handleMappingConfirm}
              onMappingReject={handleMappingReject}
              selectedValueId={selectedValueId}
            />
          )}
          
          {validationMode === 'audit' && (
            <AuditPreparationPanel
              mappingSuggestions={mappingSuggestions}
              extractedValues={extractedValues}
              onStartAudit={startAdvancedAudit}
            />
          )}
        </div>
      </div>
    </div>
  );
};

// Component for extraction summary
const ExtractionSummary = ({ extractedValues }) => {
  const stats = {
    total: extractedValues.length,
    currency: extractedValues.filter(v => v.data_type === 'currency').length,
    percentage: extractedValues.filter(v => v.data_type === 'percentage').length,
    highConfidence: extractedValues.filter(v => v.confidence >= 0.8).length
  };

  return (
    <div className="bg-blue-50 rounded-lg p-4">
      <h3 className="text-sm font-medium text-blue-900 mb-2">Extraction Summary</h3>
      <div className="grid grid-cols-2 gap-4 text-sm">
        <div>
          <span className="text-blue-600">Total Values:</span>
          <span className="font-medium text-blue-900 ml-2">{stats.total}</span>
        </div>
        <div>
          <span className="text-blue-600">High Confidence:</span>
          <span className="font-medium text-blue-900 ml-2">{stats.highConfidence}</span>
        </div>
        <div>
          <span className="text-blue-600">Currency:</span>
          <span className="font-medium text-blue-900 ml-2">{stats.currency}</span>
        </div>
        <div>
          <span className="text-blue-600">Percentages:</span>
          <span className="font-medium text-blue-900 ml-2">{stats.percentage}</span>
        </div>
      </div>
    </div>
  );
};

// Component for value review
const ValueReviewPanel = ({ extractedValues, selectedValueId, onValueSelect, onValueEdit }) => {
  const selectedValue = extractedValues.find(v => v.id === selectedValueId);

  return (
    <div className="bg-white border rounded-lg">
      <div className="p-4 border-b">
        <h3 className="text-lg font-medium">Extracted Values Review</h3>
        <p className="text-sm text-gray-600">Click on values in the document to review and edit</p>
      </div>
      
      <div className="max-h-96 overflow-y-auto">
        {extractedValues.map((value) => (
          <div
            key={value.id}
            onClick={() => onValueSelect(value.id)}
            className={`p-4 border-b cursor-pointer transition-colors ${
              selectedValueId === value.id ? 'bg-blue-50 border-blue-200' : 'hover:bg-gray-50'
            }`}
          >
            <div className="flex items-start justify-between">
              <div className="flex-1">
                <div className="flex items-center space-x-2">
                  <span className="font-medium text-gray-900">{value.value}</span>
                  <span className={`text-xs px-2 py-1 rounded-full ${
                    value.confidence >= 0.8 ? 'bg-green-100 text-green-800' :
                    value.confidence >= 0.6 ? 'bg-yellow-100 text-yellow-800' :
                    'bg-red-100 text-red-800'
                  }`}>
                    {Math.round(value.confidence * 100)}%
                  </span>
                </div>
                <p className="text-sm text-gray-600 mt-1">{value.business_context}</p>
                <p className="text-xs text-gray-500">Page {value.page_number} • {value.data_type}</p>
              </div>
              
              <button
                onClick={(e) => {
                  e.stopPropagation();
                  // Trigger edit modal
                }}
                className="p-1 text-gray-400 hover:text-gray-600"
              >
                <Edit3 className="h-4 w-4" />
              </button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
};

// Component for mapping validation
const MappingValidationPanel = ({ mappingSuggestions, onMappingConfirm, onMappingReject, selectedValueId }) => {
  return (
    <div className="bg-white border rounded-lg">
      <div className="p-4 border-b">
        <h3 className="text-lg font-medium">Mapping Suggestions</h3>
        <p className="text-sm text-gray-600">Review AI-suggested mappings to Excel sources</p>
      </div>
      
      <div className="max-h-96 overflow-y-auto">
        {mappingSuggestions.map((mapping) => (
          <MappingCard
            key={mapping.mapping_id}
            mapping={mapping}
            onConfirm={onMappingConfirm}
            onReject={onMappingReject}
            isHighlighted={selectedValueId === mapping.pdf_value_id}
          />
        ))}
      </div>
    </div>
  );
};

// Component for audit preparation
const AuditPreparationPanel = ({ mappingSuggestions, extractedValues, onStartAudit }) => {
  const confirmedMappings = mappingSuggestions.filter(m => m.user_confirmed);
  const readyForAudit = confirmedMappings.length > 0;

  return (
    <div className="bg-white border rounded-lg">
      <div className="p-4 border-b">
        <h3 className="text-lg font-medium">Audit Preparation</h3>
        <p className="text-sm text-gray-600">Review confirmation status before starting audit</p>
      </div>
      
      <div className="p-4 space-y-4">
        <div className="grid grid-cols-2 gap-4">
          <div className="bg-blue-50 p-3 rounded">
            <div className="text-2xl font-bold text-blue-900">{extractedValues.length}</div>
            <div className="text-sm text-blue-600">Total Values</div>
          </div>
          <div className="bg-green-50 p-3 rounded">
            <div className="text-2xl font-bold text-green-900">{confirmedMappings.length}</div>
            <div className="text-sm text-green-600">Confirmed Mappings</div>
          </div>
        </div>
        
        <div className="space-y-2">
          <h4 className="text-sm font-medium text-gray-900">Audit Readiness</h4>
          <div className="flex items-center space-x-2">
            {readyForAudit ? (
              <Check className="h-5 w-5 text-green-500" />
            ) : (
              <AlertTriangle className="h-5 w-5 text-yellow-500" />
            )}
            <span className="text-sm text-gray-600">
              {readyForAudit 
                ? 'Ready to start comprehensive audit'
                : 'Confirm at least one mapping to proceed'
              }
            </span>
          </div>
        </div>
        
        <button
          onClick={onStartAudit}
          disabled={!readyForAudit}
          className="w-full flex items-center justify-center px-4 py-2 text-sm font-medium text-white bg-blue-600 border border-transparent rounded-md hover:bg-blue-700 disabled:opacity-50 disabled:cursor-not-allowed"
        >
          <TrendingUp className="h-4 w-4 mr-2" />
          Start Advanced Audit
        </button>
      </div>
    </div>
  );
};

// Enhanced mapping card component
const MappingCard = ({ mapping, onConfirm, onReject, isHighlighted }) => {
  const [showDetails, setShowDetails] = useState(false);

  return (
    <div className={`p-4 border-b transition-colors ${
      isHighlighted ? 'bg-blue-50' : 'hover:bg-gray-50'
    }`}>
      <div className="flex items-start justify-between">
        <div className="flex-1">
          <div className="flex items-center space-x-2 mb-2">
            <span className="font-medium text-blue-600">{mapping.pdf_value}</span>
            <ArrowRight className="h-4 w-4 text-gray-400" />
            <span className="font-medium text-green-600">{mapping.excel_source.value}</span>
          </div>
          
          <div className="text-sm text-gray-600 space-y-1">
            <p><strong>Context:</strong> {mapping.pdf_context}</p>
            <p><strong>Source:</strong> {mapping.excel_source.sheet} • {mapping.excel_source.cell}</p>
          </div>
          
          {showDetails && (
            <div className="mt-2 p-2 bg-gray-50 rounded text-xs text-gray-600">
              <p><strong>AI Reasoning:</strong> {mapping.match_reasoning}</p>
              <p><strong>Match Type:</strong> {mapping.match_type}</p>
            </div>
          )}
        </div>
        
        <div className="flex flex-col items-end space-y-2">
          <span className={`text-xs px-2 py-1 rounded-full ${
            mapping.confidence >= 0.8 ? 'bg-green-100 text-green-800' :
            mapping.confidence >= 0.6 ? 'bg-yellow-100 text-yellow-800' :
            'bg-red-100 text-red-800'
          }`}>
            {Math.round(mapping.confidence * 100)}%
          </span>
          
          {!mapping.user_confirmed && (
            <div className="flex space-x-1">
              <button
                onClick={() => onConfirm(mapping.mapping_id)}
                className="p-1 text-green-600 hover:bg-green-100 rounded"
                title="Confirm mapping"
              >
                <Check className="h-4 w-4" />
              </button>
              <button
                onClick={() => onReject(mapping.mapping_id)}
                className="p-1 text-red-600 hover:bg-red-100 rounded"
                title="Reject mapping"
              >
                <X className="h-4 w-4" />
              </button>
            </div>
          )}
          
          <button
            onClick={() => setShowDetails(!showDetails)}
            className="p-1 text-gray-400 hover:text-gray-600"
          >
            <Eye className="h-4 w-4" />
          </button>
        </div>
      </div>
    </div>
  );
};

export default EnhancedMappingValidation;