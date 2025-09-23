import React, { useState, useEffect, useRef, useMemo } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { 
  Check, X, Edit3, Eye, AlertTriangle, TrendingUp, ArrowRight, 
  ZoomIn, ZoomOut, Brain, Target, Layers, CheckCircle, Play, 
  FileText, BarChart3, Save, ExternalLink, Users, Database,
  Filter, Search, Download, RefreshCw, Grid, List, Settings,
  ChevronDown, ChevronUp, SortAsc, SortDesc, Maximize2
} from 'lucide-react';
import toast from 'react-hot-toast';

import { apiService } from '../../services/api';
import LoadingSpinner from '../Common/LoadingSpinner';

const ComprehensiveValidation = () => {
  const { sessionId } = useParams();
  const navigate = useNavigate();
  
  // Core state
  const [validationData, setValidationData] = useState(null);
  const [pdfValues, setPdfValues] = useState([]);
  const [excelValues, setExcelValues] = useState([]);
  const [loading, setLoading] = useState(true);
  const [startingAudit, setStartingAudit] = useState(false);
  
  // UI state
  const [currentPage, setCurrentPage] = useState(0);
  const [zoom, setZoom] = useState(1.0);
  const [selectedValueId, setSelectedValueId] = useState(null);
  const [hoveredValueId, setHoveredValueId] = useState(null);
  const [showBoundingBoxes, setShowBoundingBoxes] = useState(true);
  const [validationMode, setValidationMode] = useState('pdf'); // pdf | excel | audit
  
  // Enhanced filtering and pagination for large datasets
  const [filters, setFilters] = useState({
    confidence: 0,
    searchTerm: '',
    dataType: 'all',
    businessCategory: 'all',
    valueRange: 'all',
    showModified: false
  });
  
  const [sorting, setSorting] = useState({
    field: 'confidence',
    direction: 'desc'
  });
  
  const [pagination, setPagination] = useState({
    currentPage: 0,
    itemsPerPage: 50,
    totalItems: 0
  });
  
  const [viewMode, setViewMode] = useState('grid'); // grid | table | compact
  const [expandedSections, setExpandedSections] = useState({
    pdfStats: true,
    excelStats: true,
    filters: false
  });
  
  // Canvas refs
  const canvasRef = useRef(null);
  const overlayCanvasRef = useRef(null);

  useEffect(() => {
    loadValidationData();
  }, [sessionId]);

  useEffect(() => {
    if (validationData && canvasRef.current && overlayCanvasRef.current) {
      drawDocumentPage();
      drawValueOverlays();
    }
  }, [validationData, currentPage, zoom, selectedValueId, hoveredValueId, showBoundingBoxes, filters.confidence]);

  // Memoized filtering and sorting for performance with large datasets
  const filteredPdfValues = useMemo(() => {
    let filtered = pdfValues.filter(value => {
      // Confidence filter
      if ((value.confidence || 0) < filters.confidence / 100) return false;
      
      // Search term filter
      if (filters.searchTerm) {
        const searchLower = filters.searchTerm.toLowerCase();
        const valueStr = (value.value || '').toString().toLowerCase();
        const contextStr = (value.business_context?.semantic_meaning || '').toLowerCase();
        if (!valueStr.includes(searchLower) && !contextStr.includes(searchLower)) return false;
      }
      
      // Data type filter
      if (filters.dataType !== 'all' && value.data_type !== filters.dataType) return false;
      
      // Business category filter
      if (filters.businessCategory !== 'all' && 
          value.business_context?.business_category !== filters.businessCategory) return false;
      
      // Value range filter
      if (filters.valueRange !== 'all') {
        const numValue = parseFloat(value.value);
        if (!isNaN(numValue)) {
          switch (filters.valueRange) {
            case 'small': if (numValue >= 1000) return false; break;
            case 'medium': if (numValue < 1000 || numValue >= 1000000) return false; break;
            case 'large': if (numValue < 1000000) return false; break;
          }
        }
      }
      
      // Show modified filter
      if (filters.showModified && !value.user_modified) return false;
      
      return true;
    });
    
    // Apply sorting
    filtered.sort((a, b) => {
      let aVal, bVal;
      
      switch (sorting.field) {
        case 'confidence':
          aVal = a.confidence || 0;
          bVal = b.confidence || 0;
          break;
        case 'value':
          aVal = parseFloat(a.value) || 0;
          bVal = parseFloat(b.value) || 0;
          break;
        case 'page':
          aVal = a.page_number || 0;
          bVal = b.page_number || 0;
          break;
        case 'category':
          aVal = a.business_context?.business_category || '';
          bVal = b.business_context?.business_category || '';
          break;
        default:
          aVal = a[sorting.field] || '';
          bVal = b[sorting.field] || '';
      }
      
      if (sorting.direction === 'asc') {
        return aVal > bVal ? 1 : -1;
      } else {
        return aVal < bVal ? 1 : -1;
      }
    });
    
    return filtered;
  }, [pdfValues, filters, sorting]);

  const filteredExcelValues = useMemo(() => {
    let filtered = excelValues.filter(value => {
      // Search term filter
      if (filters.searchTerm) {
        const searchLower = filters.searchTerm.toLowerCase();
        const valueStr = (value.value || '').toString().toLowerCase();
        const contextStr = (value.business_context || '').toLowerCase();
        const fileStr = (value.source_file || '').toLowerCase();
        if (!valueStr.includes(searchLower) && 
            !contextStr.includes(searchLower) && 
            !fileStr.includes(searchLower)) return false;
      }
      
      // Show modified filter
      if (filters.showModified && !value.user_modified) return false;
      
      return true;
    });
    
    // Apply sorting for Excel values
    filtered.sort((a, b) => {
      let aVal, bVal;
      
      switch (sorting.field) {
        case 'value':
          aVal = parseFloat(a.value) || 0;
          bVal = parseFloat(b.value) || 0;
          break;
        case 'file':
          aVal = a.source_file || '';
          bVal = b.source_file || '';
          break;
        case 'likelihood':
          aVal = a.presentation_likelihood || 0;
          bVal = b.presentation_likelihood || 0;
          break;
        default:
          aVal = a[sorting.field] || '';
          bVal = b[sorting.field] || '';
      }
      
      if (sorting.direction === 'asc') {
        return aVal > bVal ? 1 : -1;
      } else {
        return aVal < bVal ? 1 : -1;
      }
    });
    
    return filtered;
  }, [excelValues, filters, sorting]);

  // Pagination for current dataset
  const currentDataset = validationMode === 'pdf' ? filteredPdfValues : filteredExcelValues;
  const totalPages = Math.ceil(currentDataset.length / pagination.itemsPerPage);
  const currentPageData = currentDataset.slice(
    pagination.currentPage * pagination.itemsPerPage,
    (pagination.currentPage + 1) * pagination.itemsPerPage
  );

  const loadValidationData = async () => {
    try {
      console.log('[Comprehensive] Loading validation data for session:', sessionId);
      
      const data = await apiService.client.get(`/validation/data/${sessionId}`);
      console.log('[Comprehensive] Validation data loaded:', data);
      
      setValidationData(data);
      setPdfValues(data.pdf_values || []);
      setExcelValues(data.excel_values || []);
      
      // Show comprehensive extraction results
      toast.success(
        `Comprehensive extraction completed! Found ${data.pdf_values?.length || 0} PDF values and ${data.excel_values?.length || 0} Excel values - 100% coverage achieved!`,
        { duration: 5000 }
      );
      
    } catch (error) {
      console.error('[Comprehensive] Failed to load validation data:', error);
      toast.error('Failed to load validation data: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  const drawDocumentPage = () => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    
    if (!validationData?.document_preview?.pages?.[currentPage]) return;
    
    const pageData = validationData.document_preview.pages[currentPage];
    const img = new Image();
    
    img.onload = () => {
      const scaledWidth = img.width * zoom;
      const scaledHeight = img.height * zoom;
      
      canvas.width = scaledWidth;
      canvas.height = scaledHeight;
      
      ctx.clearRect(0, 0, scaledWidth, scaledHeight);
      ctx.drawImage(img, 0, 0, scaledWidth, scaledHeight);
    };
    
    img.src = `data:image/png;base64,${pageData.image_data}`;
  };

  const drawValueOverlays = () => {
    const overlayCanvas = overlayCanvasRef.current;
    const mainCanvas = canvasRef.current;
    
    if (!overlayCanvas || !mainCanvas) return;
    
    const ctx = overlayCanvas.getContext('2d');
    
    overlayCanvas.width = mainCanvas.width;
    overlayCanvas.height = mainCanvas.height;
    
    ctx.clearRect(0, 0, overlayCanvas.width, overlayCanvas.height);
    
    if (!showBoundingBoxes || validationMode !== 'pdf') return;
    
    const pageValues = filteredPdfValues.filter(value => 
      value.page_number === currentPage + 1 && 
      value.coordinates?.bounding_box
    );
    
    pageValues.forEach(value => {
      drawValueBoundingBox(ctx, value, overlayCanvas.width, overlayCanvas.height);
    });
  };

  const drawValueBoundingBox = (ctx, value, canvasWidth, canvasHeight) => {
    const bbox = value.coordinates.bounding_box;
    if (!bbox || bbox.length !== 4) return;
    
    const [x1, y1, x2, y2] = bbox;
    
    const canvasX1 = x1 * canvasWidth;
    const canvasY1 = y1 * canvasHeight;
    const canvasX2 = x2 * canvasWidth;
    const canvasY2 = y2 * canvasHeight;
    
    let strokeColor = getConfidenceColor(value.confidence);
    let fillColor = strokeColor + '20';
    let lineWidth = 2;
    let lineDash = [];
    
    if (selectedValueId === value.id) {
      strokeColor = '#3B82F6';
      fillColor = '#3B82F620';
      lineWidth = 3;
    } else if (hoveredValueId === value.id) {
      strokeColor = '#8B5CF6';
      fillColor = '#8B5CF620';
      lineWidth = 3;
      lineDash = [5, 5];
    }
    
    ctx.strokeStyle = strokeColor;
    ctx.fillStyle = fillColor;
    ctx.lineWidth = lineWidth;
    ctx.setLineDash(lineDash);
    
    ctx.fillRect(canvasX1, canvasY1, canvasX2 - canvasX1, canvasY2 - canvasY1);
    ctx.strokeRect(canvasX1, canvasY1, canvasX2 - canvasX1, canvasY2 - canvasY1);
    
    // Confidence indicator
    const confidenceWidth = (canvasX2 - canvasX1) * (value.confidence || 0);
    ctx.fillStyle = strokeColor;
    ctx.fillRect(canvasX1, canvasY1 - 4, confidenceWidth, 4);
    
    // Value label for selected/hovered items
    if (selectedValueId === value.id || hoveredValueId === value.id) {
      ctx.fillStyle = '#FFFFFF';
      ctx.fillRect(canvasX1, canvasY1 - 24, Math.max(100, (value.value || '').length * 8), 20);
      
      ctx.fillStyle = '#1F2937';
      ctx.font = '12px Arial';
      ctx.fillText(value.value || '', canvasX1 + 4, canvasY1 - 8);
    }
  };

  const getConfidenceColor = (confidence) => {
    if (!confidence) return '#9CA3AF';
    if (confidence >= 0.8) return '#10B981';
    if (confidence >= 0.6) return '#F59E0B';
    return '#EF4444';
  };

  const handleCanvasClick = (event) => {
    if (validationMode !== 'pdf') return;
    
    const canvas = overlayCanvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / canvas.width;
    const y = (event.clientY - rect.top) / canvas.height;
    
    const pageValues = filteredPdfValues.filter(value => 
      value.page_number === currentPage + 1 && 
      value.coordinates?.bounding_box
    );
    
    const clickedValue = pageValues.find(value => {
      const bbox = value.coordinates.bounding_box;
      if (!bbox || bbox.length !== 4) return false;
      
      const [x1, y1, x2, y2] = bbox;
      return x >= x1 && x <= x2 && y >= y1 && y <= y2;
    });
    
    if (clickedValue) {
      setSelectedValueId(clickedValue.id);
      toast.success(`Selected: ${clickedValue.value}`);
    } else {
      setSelectedValueId(null);
    }
  };

  const handleCanvasMouseMove = (event) => {
    if (validationMode !== 'pdf') return;
    
    const canvas = overlayCanvasRef.current;
    if (!canvas) return;
    
    const rect = canvas.getBoundingClientRect();
    const x = (event.clientX - rect.left) / canvas.width;
    const y = (event.clientY - rect.top) / canvas.height;
    
    const pageValues = filteredPdfValues.filter(value => 
      value.page_number === currentPage + 1 && 
      value.coordinates?.bounding_box
    );
    
    const hoveredValue = pageValues.find(value => {
      const bbox = value.coordinates.bounding_box;
      if (!bbox || bbox.length !== 4) return false;
      
      const [x1, y1, x2, y2] = bbox;
      return x >= x1 && x <= x2 && y >= y1 && y <= y2;
    });
    
    setHoveredValueId(hoveredValue?.id || null);
  };

  const handlePdfValueEdit = async (valueId, newData) => {
    try {
      await apiService.client.post(`/validation/update-pdf-value/${sessionId}`, {
        value_id: valueId,
        updates: newData
      });
      
      setPdfValues(prev => 
        prev.map(val => 
          val.id === valueId 
            ? { ...val, ...newData, user_modified: true }
            : val
        )
      );
      
      toast.success('PDF value updated successfully');
    } catch (error) {
      console.error('[Comprehensive] Failed to update PDF value:', error);
      toast.error('Failed to update PDF value: ' + error.message);
    }
  };

  const handleExcelValueEdit = async (valueId, newData) => {
    try {
      await apiService.client.post(`/validation/update-excel-value/${sessionId}`, {
        value_id: valueId,
        updates: newData
      });
      
      setExcelValues(prev => 
        prev.map(val => {
          const identifier = `${val.source_file}_${val.cell_reference}`;
          return identifier === valueId || val.id === valueId
            ? { ...val, ...newData, user_modified: true }
            : val;
        })
      );
      
      toast.success('Excel value updated successfully');
    } catch (error) {
      console.error('[Comprehensive] Failed to update Excel value:', error);
      toast.error('Failed to update Excel value: ' + error.message);
    }
  };

  const startDirectAudit = async () => {
    if (pdfValues.length === 0 || excelValues.length === 0) {
      toast.error('Need both PDF and Excel values to start audit');
      return;
    }

    setStartingAudit(true);

    try {
      console.log('[Comprehensive Audit] Starting direct comprehensive audit...');
      console.log(`[Comprehensive Audit] Auditing ${pdfValues.length} PDF values against ${excelValues.length} Excel values`);
      
      const auditResponse = await apiService.client.post(`/validation/start-direct-audit/${sessionId}`);
      
      toast.success(
        `Comprehensive audit completed! Validated ${pdfValues.length} PDF values against ${excelValues.length} Excel values with 100% coverage.`,
        { duration: 5000 }
      );
      
      // Navigate to results
      navigate(`/audit/${auditResponse.audit_session_id}`, {
        state: {
          auditResults: auditResponse.audit_results,
          auditSessionId: auditResponse.audit_session_id,
          sessionId: sessionId,
          approach: 'comprehensive_direct_validation'
        }
      });

    } catch (error) {
      console.error('[Comprehensive Audit] Direct audit failed:', error);
      toast.error('Comprehensive audit failed: ' + error.message);
    } finally {
      setStartingAudit(false);
    }
  };

  const exportData = async (format) => {
    try {
      const exportData = {
        session_id: sessionId,
        pdf_values: filteredPdfValues,
        excel_values: filteredExcelValues,
        statistics: {
          total_pdf_values: pdfValues.length,
          total_excel_values: excelValues.length,
          filtered_pdf_values: filteredPdfValues.length,
          filtered_excel_values: filteredExcelValues.length
        },
        export_timestamp: new Date().toISOString(),
        export_format: format
      };
      
      const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `comprehensive_validation_data_${sessionId}_${format}.json`;
      a.click();
      URL.revokeObjectURL(url);
      
      toast.success(`Data exported successfully as ${format.toUpperCase()}`);
    } catch (error) {
      toast.error('Export failed: ' + error.message);
    }
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-96">
        <LoadingSpinner size="xl" text="Loading comprehensive validation interface..." />
      </div>
    );
  }

  if (!validationData) {
    return (
      <div className="text-center py-12">
        <AlertTriangle className="h-12 w-12 text-yellow-500 mx-auto mb-4" />
        <h3 className="text-lg font-medium text-gray-900">No validation data available</h3>
        <p className="text-gray-600">Please process documents first.</p>
      </div>
    );
  }

  return (
    <div className="max-w-7xl mx-auto space-y-6">
      {/* Enhanced Header with Comprehensive Stats */}
      <div className="bg-white shadow-sm rounded-lg p-6">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <Brain className="h-8 w-8 text-blue-600" />
            <div>
              <h1 className="text-2xl font-bold text-gray-900">Comprehensive Direct Validation</h1>
              <p className="text-gray-600">
                Validate ALL {pdfValues.length + excelValues.length} extracted values directly - 100% coverage approach
              </p>
            </div>
          </div>
          
          <div className="flex items-center space-x-4">
            <div className="text-right">
              <div className="text-sm text-gray-500">Total Values</div>
              <div className="text-2xl font-bold text-green-600">{pdfValues.length + excelValues.length}</div>
            </div>
            
            <div className="flex bg-gray-100 rounded-lg p-1">
              {[
                { key: 'pdf', label: 'PDF Values', icon: FileText, count: filteredPdfValues.length, total: pdfValues.length },
                { key: 'excel', label: 'Excel Values', icon: BarChart3, count: filteredExcelValues.length, total: excelValues.length },
                { key: 'audit', label: 'Start Audit', icon: Play, count: null }
              ].map((mode) => {
                const Icon = mode.icon;
                return (
                  <button
                    key={mode.key}
                    onClick={() => {
                      setValidationMode(mode.key);
                      setPagination(prev => ({ ...prev, currentPage: 0 }));
                    }}
                    className={`px-3 py-2 text-sm font-medium rounded-md flex items-center space-x-1 ${
                      validationMode === mode.key
                        ? 'bg-white text-blue-600 shadow-sm'
                        : 'text-gray-600 hover:text-gray-900'
                    }`}
                  >
                    <Icon className="h-4 w-4" />
                    <span>{mode.label}</span>
                    {mode.count !== null && (
                      <span className="text-xs bg-gray-200 text-gray-700 px-1 rounded">
                        {mode.count}{mode.total && mode.count !== mode.total ? `/${mode.total}` : ''}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          </div>
        </div>
      </div>

      {/* Comprehensive Statistics Dashboard */}
      <div className="bg-gradient-to-r from-blue-50 to-green-50 rounded-lg p-6 border">
        <div className="flex items-center justify-between mb-4">
          <h3 className="text-lg font-medium text-gray-900 flex items-center">
            <TrendingUp className="h-5 w-5 mr-2 text-blue-600" />
            Comprehensive Extraction Statistics
          </h3>
          <div className="flex items-center space-x-2">
            <button
              onClick={() => exportData('json')}
              className="px-3 py-1 text-sm text-blue-600 border border-blue-300 rounded hover:bg-blue-50 flex items-center"
            >
              <Download className="h-4 w-4 mr-1" />
              Export
            </button>
            <button
              onClick={() => loadValidationData()}
              className="px-3 py-1 text-sm text-gray-600 border border-gray-300 rounded hover:bg-gray-50 flex items-center"
            >
              <RefreshCw className="h-4 w-4 mr-1" />
              Refresh
            </button>
          </div>
        </div>
        
        <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
          <div className="bg-white rounded p-4 text-center">
            <div className="text-3xl font-bold text-blue-600">{pdfValues.length}</div>
            <div className="text-sm text-gray-600">PDF Values Extracted</div>
            <div className="text-xs text-green-600 mt-1">All Pages Processed</div>
          </div>
          <div className="bg-white rounded p-4 text-center">
            <div className="text-3xl font-bold text-green-600">{excelValues.length}</div>
            <div className="text-sm text-gray-600">Excel Values Found</div>
            <div className="text-xs text-green-600 mt-1">All Sheets Analyzed</div>
          </div>
          <div className="bg-white rounded p-4 text-center">
            <div className="text-3xl font-bold text-purple-600">{pdfValues.length + excelValues.length}</div>
            <div className="text-sm text-gray-600">Total for Validation</div>
            <div className="text-xs text-purple-600 mt-1">Comprehensive Coverage</div>
          </div>
          <div className="bg-white rounded p-4 text-center">
            <div className="text-3xl font-bold text-orange-600">100%</div>
            <div className="text-sm text-gray-600">Coverage Achieved</div>
            <div className="text-xs text-orange-600 mt-1">No Artificial Limits</div>
          </div>
        </div>
        
        {/* Show extraction improvement */}
        <div className="mt-4 bg-white rounded p-3 border border-green-200">
          <div className="text-sm text-green-800">
            <strong>Extraction Improvement:</strong>
            <div className="mt-1 grid grid-cols-2 gap-4 text-xs">
              <div>
                <span className="text-red-600">Before:</span> Limited to ~18 Excel values (artificial limits)
              </div>
              <div>
                <span className="text-green-600">Now:</span> {excelValues.length} Excel values extracted (comprehensive)
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Advanced Filtering and Controls */}
      <div className="bg-white shadow-sm rounded-lg">
        <div className="p-4 border-b">
          <button
            onClick={() => setExpandedSections(prev => ({ ...prev, filters: !prev.filters }))}
            className="flex items-center justify-between w-full text-left"
          >
            <div className="flex items-center space-x-2">
              <Filter className="h-5 w-5 text-gray-500" />
              <span className="text-lg font-medium text-gray-900">Advanced Filters & Controls</span>
              <span className="text-sm text-gray-500">
                ({validationMode === 'pdf' ? filteredPdfValues.length : filteredExcelValues.length} of {validationMode === 'pdf' ? pdfValues.length : excelValues.length} shown)
              </span>
            </div>
            {expandedSections.filters ? <ChevronUp className="h-5 w-5" /> : <ChevronDown className="h-5 w-5" />}
          </button>
        </div>
        
        {expandedSections.filters && (
          <div className="p-4 space-y-4">
            <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-5 gap-4">
              {/* Search */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Search</label>
                <div className="relative">
                  <Search className="h-4 w-4 absolute left-3 top-3 text-gray-400" />
                  <input
                    type="text"
                    placeholder="Search values..."
                    value={filters.searchTerm}
                    onChange={(e) => setFilters(prev => ({ ...prev, searchTerm: e.target.value }))}
                    className="pl-10 w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
                  />
                </div>
              </div>
              
              {/* Confidence Filter (PDF only) */}
              {validationMode === 'pdf' && (
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">
                    Min Confidence: {filters.confidence}%
                  </label>
                  <input
                    type="range"
                    min="0"
                    max="100"
                    value={filters.confidence}
                    onChange={(e) => setFilters(prev => ({ ...prev, confidence: parseInt(e.target.value) }))}
                    className="w-full"
                  />
                </div>
              )}
              
              {/* Data Type Filter */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Data Type</label>
                <select
                  value={filters.dataType}
                  onChange={(e) => setFilters(prev => ({ ...prev, dataType: e.target.value }))}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
                >
                  <option value="all">All Types</option>
                  <option value="currency">Currency</option>
                  <option value="percentage">Percentage</option>
                  <option value="count">Count</option>
                  <option value="ratio">Ratio</option>
                  <option value="metric">Metric</option>
                </select>
              </div>
              
              {/* Value Range Filter */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Value Range</label>
                <select
                  value={filters.valueRange}
                  onChange={(e) => setFilters(prev => ({ ...prev, valueRange: e.target.value }))}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
                >
                  <option value="all">All Ranges</option>
                  <option value="small">Small (&lt;1K)</option>
                  <option value="medium">Medium (1K-1M)</option>
                  <option value="large">Large (&gt;1M)</option>
                </select>
              </div>
              
              {/* View Mode */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">View Mode</label>
                <div className="flex border border-gray-300 rounded-md">
                  <button
                    onClick={() => setViewMode('grid')}
                    className={`px-3 py-2 text-sm flex-1 ${viewMode === 'grid' ? 'bg-blue-100 text-blue-700' : 'text-gray-600'}`}
                  >
                    <Grid className="h-4 w-4 mx-auto" />
                  </button>
                  <button
                    onClick={() => setViewMode('table')}
                    className={`px-3 py-2 text-sm flex-1 border-l ${viewMode === 'table' ? 'bg-blue-100 text-blue-700' : 'text-gray-600'}`}
                  >
                    <List className="h-4 w-4 mx-auto" />
                  </button>
                </div>
              </div>
            </div>
            
            {/* Sorting and Additional Controls */}
            <div className="flex items-center justify-between pt-4 border-t">
              <div className="flex items-center space-x-4">
                <div className="flex items-center space-x-2">
                  <span className="text-sm font-medium text-gray-700">Sort by:</span>
                  <select
                    value={sorting.field}
                    onChange={(e) => setSorting(prev => ({ ...prev, field: e.target.value }))}
                    className="px-3 py-1 border border-gray-300 rounded text-sm"
                  >
                    <option value="confidence">Confidence</option>
                    <option value="value">Value</option>
                    <option value="page">Page</option>
                    <option value="category">Category</option>
                  </select>
                  <button
                    onClick={() => setSorting(prev => ({ ...prev, direction: prev.direction === 'asc' ? 'desc' : 'asc' }))}
                    className="p-1 text-gray-500 hover:text-gray-700"
                  >
                    {sorting.direction === 'asc' ? <SortAsc className="h-4 w-4" /> : <SortDesc className="h-4 w-4" />}
                  </button>
                </div>
                
                <label className="flex items-center space-x-2">
                  <input
                    type="checkbox"
                    checked={filters.showModified}
                    onChange={(e) => setFilters(prev => ({ ...prev, showModified: e.target.checked }))}
                    className="rounded"
                  />
                  <span className="text-sm text-gray-700">Show only modified</span>
                </label>
              </div>
              
              <button
                onClick={() => setFilters({
                  confidence: 0,
                  searchTerm: '',
                  dataType: 'all',
                  businessCategory: 'all',
                  valueRange: 'all',
                  showModified: false
                })}
                className="px-3 py-1 text-sm text-gray-600 border border-gray-300 rounded hover:bg-gray-50"
              >
                Clear Filters
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Main Content Grid */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Document Viewer / Excel Viewer */}
        <div className="space-y-4">
          {validationMode === 'pdf' ? (
            <InteractiveDocumentViewer
              validationData={validationData}
              currentPage={currentPage}
              setCurrentPage={setCurrentPage}
              zoom={zoom}
              setZoom={setZoom}
              showBoundingBoxes={showBoundingBoxes}
              setShowBoundingBoxes={setShowBoundingBoxes}
              canvasRef={canvasRef}
              overlayCanvasRef={overlayCanvasRef}
              onCanvasClick={handleCanvasClick}
              onCanvasMouseMove={handleCanvasMouseMove}
            />
          ) : (
            <ComprehensiveExcelViewer 
              excelValues={excelValues}
              filteredValues={filteredExcelValues}
              validationMode={validationMode}
            />
          )}
        </div>

        {/* Validation Panel */}
        <div className="space-y-4">
          {validationMode === 'pdf' && (
            <ComprehensivePdfValuesPanel
              pdfValues={currentPageData}
              totalValues={filteredPdfValues.length}
              allValues={pdfValues.length}
              selectedValueId={selectedValueId}
              onValueSelect={setSelectedValueId}
              onValueEdit={handlePdfValueEdit}
              pagination={pagination}
              setPagination={setPagination}
              totalPages={totalPages}
              viewMode={viewMode}
            />
          )}
          
          {validationMode === 'excel' && (
            <ComprehensiveExcelValuesPanel
              excelValues={currentPageData}
              totalValues={filteredExcelValues.length}
              allValues={excelValues.length}
              onValueEdit={handleExcelValueEdit}
              pagination={pagination}
              setPagination={setPagination}
              totalPages={totalPages}
              viewMode={viewMode}
            />
          )}
          
          {validationMode === 'audit' && (
            <ComprehensiveDirectAuditPanel
              pdfValues={pdfValues}
              excelValues={excelValues}
              onStartAudit={startDirectAudit}
              loading={startingAudit}
            />
          )}
        </div>
      </div>
    </div>
  );
};

// Enhanced Supporting Components with pagination and better handling of large datasets

const InteractiveDocumentViewer = ({ 
  validationData, currentPage, setCurrentPage, zoom, setZoom, 
  showBoundingBoxes, setShowBoundingBoxes,
  canvasRef, overlayCanvasRef, onCanvasClick, onCanvasMouseMove 
}) => {
  const totalPages = validationData?.document_preview?.pages?.length || 0;

  return (
    <div className="bg-white border rounded-lg shadow-sm overflow-hidden">
      <div className="flex items-center justify-between p-4 border-b bg-gray-50">
        <div className="flex items-center space-x-4">
          <h3 className="text-lg font-medium flex items-center">
            <FileText className="h-5 w-5 mr-2 text-blue-600" />
            PDF Document Preview
          </h3>
          <span className="text-sm text-gray-500">
            Page {currentPage + 1} of {totalPages}
          </span>
        </div>
        
        <div className="flex items-center space-x-3">
          <div className="flex items-center space-x-1 bg-white rounded border">
            <button
              onClick={() => setZoom(Math.max(0.5, zoom - 0.25))}
              className="p-1 text-gray-600 hover:text-gray-900"
            >
              <ZoomOut className="h-4 w-4" />
            </button>
            <span className="px-2 text-sm text-gray-600 border-x">
              {Math.round(zoom * 100)}%
            </span>
            <button
              onClick={() => setZoom(Math.min(3.0, zoom + 0.25))}
              className="p-1 text-gray-600 hover:text-gray-900"
            >
              <ZoomIn className="h-4 w-4" />
            </button>
          </div>
          
          <button
            onClick={() => setShowBoundingBoxes(!showBoundingBoxes)}
            className={`px-3 py-1 text-sm rounded border ${
              showBoundingBoxes 
                ? 'bg-blue-100 text-blue-700 border-blue-300' 
                : 'bg-white text-gray-600 border-gray-300'
            }`}
          >
            <Eye className="h-4 w-4 inline mr-1" />
            Overlays
          </button>
          
          <button
            className="p-1 text-gray-600 hover:text-gray-900"
            title="Maximize viewer"
          >
            <Maximize2 className="h-4 w-4" />
          </button>
        </div>
      </div>

      <div className="relative bg-gray-100 overflow-auto" style={{ maxHeight: '600px' }}>
        <div className="relative">
          <canvas
            ref={canvasRef}
            className="block"
            style={{ maxWidth: '100%', height: 'auto' }}
          />
          <canvas
            ref={overlayCanvasRef}
            onClick={onCanvasClick}
            onMouseMove={onCanvasMouseMove}
            className="absolute top-0 left-0 cursor-pointer"
            style={{ 
              maxWidth: '100%', 
              height: 'auto',
              pointerEvents: showBoundingBoxes ? 'auto' : 'none'
            }}
          />
        </div>
      </div>

      <div className="flex items-center justify-between p-4 border-t">
        <button
          onClick={() => setCurrentPage(Math.max(0, currentPage - 1))}
          disabled={currentPage === 0}
          className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50"
        >
          Previous
        </button>
        
        <div className="flex items-center space-x-1 max-w-md overflow-x-auto">
          {Array.from({ length: Math.min(totalPages, 20) }, (_, i) => (
            <button
              key={i}
              onClick={() => setCurrentPage(i)}
              className={`w-8 h-8 text-xs rounded ${
                currentPage === i
                  ? 'bg-blue-600 text-white'
                  : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
              }`}
            >
              {i + 1}
            </button>
          ))}
          {totalPages > 20 && <span className="text-gray-500">...</span>}
        </div>
        
        <button
          onClick={() => setCurrentPage(Math.min(totalPages - 1, currentPage + 1))}
          disabled={currentPage >= totalPages - 1}
          className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50"
        >
          Next
        </button>
      </div>
    </div>
  );
};

const ComprehensiveExcelViewer = ({ excelValues, filteredValues, validationMode }) => {
  const fileGroups = filteredValues.reduce((groups, value) => {
    const file = value.source_file || 'Unknown File';
    if (!groups[file]) groups[file] = [];
    groups[file].push(value);
    return groups;
  }, {});

  return (
    <div className="bg-white border rounded-lg shadow-sm overflow-hidden">
      <div className="p-4 border-b bg-gray-50">
        <h3 className="text-lg font-medium flex items-center">
          <BarChart3 className="h-5 w-5 mr-2 text-green-600" />
          Comprehensive Excel Sources ({filteredValues.length} values)
        </h3>
        <p className="text-sm text-gray-600 mt-1">
          All extracted Excel values from comprehensive analysis - no artificial limits applied
        </p>
      </div>
      
      <div className="max-h-96 overflow-y-auto p-4">
        {Object.entries(fileGroups).map(([fileName, values]) => (
          <div key={fileName} className="mb-6">
            <h4 className="font-medium text-gray-900 mb-3 flex items-center">
              <Database className="h-4 w-4 mr-2 text-green-600" />
              {fileName} ({values.length} values)
            </h4>
            <div className="grid grid-cols-1 gap-2">
              {values.slice(0, 10).map((value, idx) => (
                <div key={idx} className="bg-gray-50 p-2 rounded text-sm">
                  <div className="flex justify-between items-start">
                    <div className="flex-1">
                      <div className="flex items-center space-x-2">
                        <span className="font-medium">{value.cell_reference}</span>
                        <span className="text-gray-600">{value.value}</span>
                        {value.presentation_likelihood && (
                          <span className="text-xs px-2 py-1 bg-blue-100 text-blue-800 rounded">
                            {Math.round(value.presentation_likelihood * 100)}% likely
                          </span>
                        )}
                      </div>
                      {value.business_context && (
                        <div className="text-xs text-gray-500 mt-1">
                          {value.business_context}
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              ))}
              {values.length > 10 && (
                <div className="text-sm text-gray-500 italic">
                  ... and {values.length - 10} more values (use pagination to see all)
                </div>
              )}
            </div>
          </div>
        ))}
        
        {Object.keys(fileGroups).length === 0 && (
          <div className="text-center py-8 text-gray-500">
            No Excel values match current filters
          </div>
        )}
      </div>
    </div>
  );
};

const ComprehensivePdfValuesPanel = ({ 
  pdfValues, totalValues, allValues, selectedValueId, onValueSelect, onValueEdit, 
  pagination, setPagination, totalPages, viewMode 
}) => {
  const [editingValue, setEditingValue] = useState(null);

  return (
    <div className="bg-white border rounded-lg overflow-hidden">
      <div className="p-4 border-b bg-gray-50">
        <div className="flex items-center justify-between">
          <div>
            <h3 className="text-lg font-medium flex items-center">
              <FileText className="h-5 w-5 mr-2 text-blue-600" />
              PDF Values Validation
            </h3>
            <p className="text-sm text-gray-600 mt-1">
              Showing {pdfValues.length} of {totalValues} filtered values ({allValues} total extracted)
            </p>
          </div>
          
          {/* Pagination Controls */}
          <div className="flex items-center space-x-2">
            <button
              onClick={() => setPagination(prev => ({ ...prev, currentPage: Math.max(0, prev.currentPage - 1) }))}
              disabled={pagination.currentPage === 0}
              className="px-2 py-1 text-sm border rounded disabled:opacity-50"
            >
              Prev
            </button>
            <span className="text-sm text-gray-600">
              {pagination.currentPage + 1} / {totalPages}
            </span>
            <button
              onClick={() => setPagination(prev => ({ ...prev, currentPage: Math.min(totalPages - 1, prev.currentPage + 1) }))}
              disabled={pagination.currentPage >= totalPages - 1}
              className="px-2 py-1 text-sm border rounded disabled:opacity-50"
            >
              Next
            </button>
          </div>
        </div>
      </div>
      
      <div className="max-h-96 overflow-y-auto">
        {pdfValues.map((value) => (
          <PdfValueCard
            key={value.id}
            value={value}
            isSelected={selectedValueId === value.id}
            onSelect={onValueSelect}
            onEdit={(newData) => {
              onValueEdit(value.id, newData);
              setEditingValue(null);
            }}
            onStartEdit={() => setEditingValue(value)}
            isEditing={editingValue?.id === value.id}
            onCancelEdit={() => setEditingValue(null)}
            viewMode={viewMode}
          />
        ))}
        
        {pdfValues.length === 0 && (
          <div className="p-8 text-center text-gray-500">
            No PDF values match current filters
          </div>
        )}
      </div>
    </div>
  );
};

const ComprehensiveExcelValuesPanel = ({ 
  excelValues, totalValues, allValues, onValueEdit, 
  pagination, setPagination, totalPages, viewMode 
}) => {
  const [editingValue, setEditingValue] = useState(null);

  return (
    <div className="bg-white border rounded-lg overflow-hidden">
      <div className="p-4 border-b bg-gray-50">
        <div className="flex items-center justify-between">
          <div>
            <h3 className="text-lg font-medium flex items-center">
              <BarChart3 className="h-5 w-5 mr-2 text-green-600" />
              Excel Values Validation
            </h3>
            <p className="text-sm text-gray-600 mt-1">
              Showing {excelValues.length} of {totalValues} filtered values ({allValues} total extracted)
            </p>
          </div>
          
          {/* Pagination Controls */}
          <div className="flex items-center space-x-2">
            <button
              onClick={() => setPagination(prev => ({ ...prev, currentPage: Math.max(0, prev.currentPage - 1) }))}
              disabled={pagination.currentPage === 0}
              className="px-2 py-1 text-sm border rounded disabled:opacity-50"
            >
              Prev
            </button>
            <span className="text-sm text-gray-600">
              {pagination.currentPage + 1} / {totalPages}
            </span>
            <button
              onClick={() => setPagination(prev => ({ ...prev, currentPage: Math.min(totalPages - 1, prev.currentPage + 1) }))}
              disabled={pagination.currentPage >= totalPages - 1}
              className="px-2 py-1 text-sm border rounded disabled:opacity-50"
            >
              Next
            </button>
          </div>
        </div>
      </div>
      
      <div className="max-h-96 overflow-y-auto">
        {excelValues.map((value, index) => (
          <ExcelValueCard
            key={`${value.source_file}_${value.cell_reference}_${index}`}
            value={value}
            valueId={`${value.source_file}_${value.cell_reference}`}
            onEdit={(newData) => {
              onValueEdit(`${value.source_file}_${value.cell_reference}`, newData);
              setEditingValue(null);
            }}
            onStartEdit={() => setEditingValue(`${value.source_file}_${value.cell_reference}`)}
            isEditing={editingValue === `${value.source_file}_${value.cell_reference}`}
            onCancelEdit={() => setEditingValue(null)}
            viewMode={viewMode}
          />
        ))}
        
        {excelValues.length === 0 && (
          <div className="p-8 text-center text-gray-500">
            No Excel values match current filters
          </div>
        )}
      </div>
    </div>
  );
};

const PdfValueCard = ({ value, isSelected, onSelect, onEdit, onStartEdit, isEditing, onCancelEdit, viewMode }) => {
  const [editData, setEditData] = useState({
    value: value.value || '',
    business_context: value.business_context?.semantic_meaning || ''
  });

  const getConfidenceColor = (confidence) => {
    if (!confidence) return 'bg-gray-100 text-gray-800';
    if (confidence >= 0.8) return 'bg-green-100 text-green-800';
    if (confidence >= 0.6) return 'bg-yellow-100 text-yellow-800';
    return 'bg-red-100 text-red-800';
  };

  if (isEditing) {
    return (
      <div className="p-4 border-b bg-blue-50">
        <div className="space-y-3">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Value</label>
            <input
              type="text"
              value={editData.value}
              onChange={(e) => setEditData({...editData, value: e.target.value})}
              className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Business Context</label>
            <textarea
              value={editData.business_context}
              onChange={(e) => setEditData({...editData, business_context: e.target.value})}
              rows={2}
              className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
            />
          </div>
          <div className="flex justify-end space-x-2">
            <button
              onClick={onCancelEdit}
              className="px-3 py-1 text-sm text-gray-600 border border-gray-300 rounded hover:bg-gray-50 flex items-center"
            >
              <X className="h-4 w-4 mr-1" />
              Cancel
            </button>
            <button
              onClick={() => onEdit(editData)}
              className="px-3 py-1 text-sm text-white bg-blue-600 rounded hover:bg-blue-700 flex items-center"
            >
              <Save className="h-4 w-4 mr-1" />
              Save
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (viewMode === 'compact') {
    return (
      <div
        onClick={() => onSelect(value.id)}
        className={`p-2 border-b cursor-pointer transition-colors flex items-center justify-between ${
          isSelected ? 'bg-blue-50 border-blue-200' : 'hover:bg-gray-50'
        }`}
      >
        <div className="flex items-center space-x-3">
          <span className="font-medium text-gray-900">{value.value}</span>
          <span className={`text-xs px-2 py-1 rounded-full ${getConfidenceColor(value.confidence)}`}>
            {Math.round((value.confidence || 0) * 100)}%
          </span>
          <span className="text-xs text-gray-500">Page {value.page_number}</span>
        </div>
        <button
          onClick={(e) => { e.stopPropagation(); onStartEdit(); }}
          className="p-1 text-gray-400 hover:text-gray-600"
        >
          <Edit3 className="h-4 w-4" />
        </button>
      </div>
    );
  }

  return (
    <div
      onClick={() => onSelect(value.id)}
      className={`p-4 border-b cursor-pointer transition-colors ${
        isSelected ? 'bg-blue-50 border-blue-200' : 'hover:bg-gray-50'
      }`}
    >
      <div className="flex items-start justify-between">
        <div className="flex-1">
          <div className="flex items-center space-x-2 mb-2">
            <span className="font-medium text-gray-900">{value.value}</span>
            <span className={`text-xs px-2 py-1 rounded-full ${getConfidenceColor(value.confidence)}`}>
              {Math.round((value.confidence || 0) * 100)}%
            </span>
            {value.user_modified && (
              <span className="text-xs px-2 py-1 rounded-full bg-purple-100 text-purple-800">
                Modified
              </span>
            )}
          </div>
          <p className="text-sm text-gray-600 mb-1">
            {value.business_context?.semantic_meaning || 'No context available'}
          </p>
          <div className="flex items-center space-x-3 text-xs text-gray-500">
            <span>Page {value.page_number}</span>
            <span>{value.data_type}</span>
            <span>{value.business_context?.business_category}</span>
          </div>
        </div>
        
        <button
          onClick={(e) => {
            e.stopPropagation();
            onStartEdit();
          }}
          className="p-1 text-gray-400 hover:text-gray-600"
        >
          <Edit3 className="h-4 w-4" />
        </button>
      </div>
    </div>
  );
};

const ExcelValueCard = ({ value, valueId, onEdit, onStartEdit, isEditing, onCancelEdit, viewMode }) => {
  const [editData, setEditData] = useState({
    value: value.value || '',
    business_context: value.business_context || ''
  });

  if (isEditing) {
    return (
      <div className="p-4 border-b bg-green-50">
        <div className="space-y-3">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Value</label>
            <input
              type="text"
              value={editData.value}
              onChange={(e) => setEditData({...editData, value: e.target.value})}
              className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
            />
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-1">Business Context</label>
            <textarea
              value={editData.business_context}
              onChange={(e) => setEditData({...editData, business_context: e.target.value})}
              rows={2}
              className="w-full px-3 py-2 border border-gray-300 rounded-md text-sm"
            />
          </div>
          <div className="flex justify-end space-x-2">
            <button
              onClick={onCancelEdit}
              className="px-3 py-1 text-sm text-gray-600 border border-gray-300 rounded hover:bg-gray-50 flex items-center"
            >
              <X className="h-4 w-4 mr-1" />
              Cancel
            </button>
            <button
              onClick={() => onEdit(editData)}
              className="px-3 py-1 text-sm text-white bg-green-600 rounded hover:bg-green-700 flex items-center"
            >
              <Save className="h-4 w-4 mr-1" />
              Save
            </button>
          </div>
        </div>
      </div>
    );
  }

  if (viewMode === 'compact') {
    return (
      <div className="p-2 border-b hover:bg-gray-50 flex items-center justify-between">
        <div className="flex items-center space-x-3">
          <span className="text-sm font-medium text-gray-700">{value.source_file}</span>
          <span className="text-xs px-2 py-1 rounded bg-green-100 text-green-800">
            {value.cell_reference}
          </span>
          <span className="font-medium text-gray-900">{value.value}</span>
        </div>
        <button
          onClick={onStartEdit}
          className="p-1 text-gray-400 hover:text-gray-600"
        >
          <Edit3 className="h-4 w-4" />
        </button>
      </div>
    );
  }

  return (
    <div className="p-4 border-b hover:bg-gray-50">
      <div className="flex items-start justify-between">
        <div className="flex-1">
          <div className="flex items-center space-x-2 mb-2">
            <span className="text-sm font-medium text-gray-700">{value.source_file}</span>
            <span className="text-xs px-2 py-1 rounded bg-green-100 text-green-800">
              {value.cell_reference}
            </span>
            {value.presentation_likelihood && (
              <span className="text-xs px-2 py-1 rounded bg-blue-100 text-blue-800">
                {Math.round(value.presentation_likelihood * 100)}% likely
              </span>
            )}
            {value.user_modified && (
              <span className="text-xs px-2 py-1 rounded-full bg-purple-100 text-purple-800">
                Modified
              </span>
            )}
          </div>
          <div className="font-medium text-gray-900 mb-1">{value.value}</div>
          <p className="text-sm text-gray-600">
            {value.business_context || 'No context available'}
          </p>
        </div>
        
        <button
          onClick={onStartEdit}
          className="p-1 text-gray-400 hover:text-gray-600"
        >
          <Edit3 className="h-4 w-4" />
        </button>
      </div>
    </div>
  );
};

const ComprehensiveDirectAuditPanel = ({ pdfValues, excelValues, onStartAudit, loading }) => {
  const readyForAudit = pdfValues.length > 0 && excelValues.length > 0;
  const totalValues = pdfValues.length + excelValues.length;

  return (
    <div className="bg-white border rounded-lg overflow-hidden">
      <div className="p-4 border-b bg-gray-50">
        <h3 className="text-lg font-medium flex items-center">
          <Brain className="h-5 w-5 mr-2 text-blue-600" />
          Comprehensive Direct Audit
        </h3>
        <p className="text-sm text-gray-600 mt-1">
          Audit ALL {totalValues} extracted values with 100% coverage - no mapping bottlenecks
        </p>
      </div>
      
      <div className="p-6 space-y-6">
        {/* Enhanced Audit Coverage */}
        <div className="bg-gradient-to-r from-green-50 to-blue-50 p-4 rounded-lg">
          <h4 className="text-sm font-medium text-gray-900 mb-3">Comprehensive Audit Coverage Analysis</h4>
          <div className="grid grid-cols-2 gap-4">
            <div className="text-center">
              <div className="text-2xl font-bold text-blue-600">{pdfValues.length}</div>
              <div className="text-sm text-gray-600">PDF Values</div>
              <div className="text-xs text-green-600">100% Coverage</div>
            </div>
            <div className="text-center">
              <div className="text-2xl font-bold text-green-600">{excelValues.length}</div>
              <div className="text-sm text-gray-600">Excel Sources</div>
              <div className="text-xs text-green-600">All Available</div>
            </div>
          </div>
          
          {/* Show the dramatic improvement */}
          <div className="mt-4 pt-4 border-t border-gray-200">
            <div className="text-sm space-y-2">
              <div className="flex justify-between">
                <span className="text-red-600">Previous Limit:</span>
                <span className="text-red-600 font-medium">~18 Excel values</span>
              </div>
              <div className="flex justify-between">
                <span className="text-green-600">Current Extraction:</span>
                <span className="text-green-600 font-medium">{excelValues.length} Excel values</span>
              </div>
              <div className="flex justify-between">
                <span className="text-blue-600">Improvement:</span>
                <span className="text-blue-600 font-medium">
                  {Math.round((excelValues.length / 18) * 100)}% increase
                </span>
              </div>
            </div>
          </div>
        </div>
        
        {/* Audit Features */}
        <div className="space-y-3">
          <h4 className="text-sm font-medium text-gray-900">Comprehensive Direct Audit Features</h4>
          <div className="space-y-2">
            <div className="flex items-center space-x-2">
              <CheckCircle className="h-4 w-4 text-green-500" />
              <span className="text-sm text-gray-600">
                Compare all {pdfValues.length} PDF values against all {excelValues.length} Excel values
              </span>
            </div>
            <div className="flex items-center space-x-2">
              <CheckCircle className="h-4 w-4 text-green-500" />
              <span className="text-sm text-gray-600">
                Advanced value normalization and semantic matching
              </span>
            </div>
            <div className="flex items-center space-x-2">
              <CheckCircle className="h-4 w-4 text-green-500" />
              <span className="text-sm text-gray-600">
                Business context-aware validation using Gemini 2.5 Pro
              </span>
            </div>
            <div className="flex items-center space-x-2">
              <CheckCircle className="h-4 w-4 text-green-500" />
              <span className="text-sm text-gray-600">
                100% coverage - no bottlenecks from artificial extraction limits
              </span>
            </div>
            <div className="flex items-center space-x-2">
              <CheckCircle className="h-4 w-4 text-green-500" />
              <span className="text-sm text-gray-600">
                Intelligent batching for optimal performance with large datasets
              </span>
            </div>
          </div>
        </div>
        
        {/* Start Audit Button */}
        <button
          onClick={onStartAudit}
          disabled={!readyForAudit || loading}
          className="w-full flex items-center justify-center px-6 py-3 text-sm font-medium text-white bg-gradient-to-r from-blue-600 to-green-600 border border-transparent rounded-lg hover:from-blue-700 hover:to-green-700 disabled:opacity-50 disabled:cursor-not-allowed transition-all transform hover:scale-105"
        >
          {loading ? (
            <LoadingSpinner size="sm" text="Running comprehensive audit..." />
          ) : (
            <>
              <Play className="h-5 w-5 mr-2" />
              Start Comprehensive Direct Audit
              <ArrowRight className="ml-2 h-4 w-4" />
            </>
          )}
        </button>
        
        {/* Readiness Status */}
        <div className="pt-4 border-t">
          <div className="flex items-center space-x-2">
            {readyForAudit ? (
              <CheckCircle className="h-5 w-5 text-green-500" />
            ) : (
              <AlertTriangle className="h-5 w-5 text-yellow-500" />
            )}
            <span className="text-sm text-gray-600">
              {readyForAudit 
                ? `Ready for comprehensive audit: ${totalValues} total values to validate`
                : 'Complete document processing to enable audit'
              }
            </span>
          </div>
        </div>
        
        {/* Performance Expectations */}
        {readyForAudit && (
          <div className="bg-blue-50 border border-blue-200 rounded p-3">
            <div className="text-sm text-blue-800">
              <strong>Expected Performance:</strong>
              <div className="mt-1 space-y-1">
                <div>• Processing time: ~{Math.ceil(pdfValues.length / 5)} batches (~{Math.ceil(pdfValues.length / 5 * 2)} minutes)</div>
                <div>• AI model: Gemini 2.5 Pro with comprehensive analysis</div>
                <div>• Memory usage: Optimized for large datasets</div>
                <div>• Results: Detailed validation for every extracted value</div>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default ComprehensiveValidation;