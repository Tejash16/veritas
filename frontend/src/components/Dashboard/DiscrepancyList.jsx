import React, { useState } from 'react';
import { AlertTriangle, CheckCircle, XCircle, Eye, ChevronDown, ChevronUp } from 'lucide-react';

const DiscrepancyList = ({ discrepancies = [] }) => {
  const [filter, setFilter] = useState('all');
  const [sortBy, setSortBy] = useState('confidence');
  const [expandedItems, setExpandedItems] = useState(new Set());

  const getStatusIcon = (status) => {
    switch (status) {
      case 'matched':
        return <CheckCircle className="h-5 w-5 text-green-500" />;
      case 'mismatched':
        return <XCircle className="h-5 w-5 text-red-500" />;
      case 'formatting_error':
        return <AlertTriangle className="h-5 w-5 text-yellow-500" />;
      default:
        return <Eye className="h-5 w-5 text-gray-500" />;
    }
  };

  const getStatusColor = (status) => {
    switch (status) {
      case 'matched':
        return 'bg-green-50 text-green-800 border-green-200';
      case 'mismatched':
        return 'bg-red-50 text-red-800 border-red-200';
      case 'formatting_error':
        return 'bg-yellow-50 text-yellow-800 border-yellow-200';
      default:
        return 'bg-gray-50 text-gray-800 border-gray-200';
    }
  };

  const toggleExpanded = (index) => {
    const newExpanded = new Set(expandedItems);
    // console.log('newExpanded', discrepancies);
    if (newExpanded.has(index)) {
      newExpanded.delete(index);
    } else {
      newExpanded.add(index);
    }
    setExpandedItems(newExpanded);
  };

  const filteredDiscrepancies = discrepancies
    .filter(item => filter === 'all' || item.validation_status === filter)
    .sort((a, b) => {
      switch (sortBy) {
        case 'confidence':
          return (b.confidence_score || 0) - (a.confidence_score || 0);
        case 'slide':
          return (a.pdf_slide || 0) - (b.pdf_slide || 0);
        case 'status':
          return a.validation_status.localeCompare(b.validation_status);
        default:
          return 0;
      }
    });

  return (
    <div className="space-y-4">
      {/* Filters */}
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div className="flex items-center space-x-4">
          <div>
            <label htmlFor="filter" className="text-sm font-medium text-gray-700 mr-2">
              Filter:
            </label>
            <select
              id="filter"
              value={filter}
              onChange={(e) => setFilter(e.target.value)}
              className="text-sm border-gray-300 rounded-md"
            >
              <option value="all">All Items</option>
              <option value="matched">Matched</option>
              <option value="mismatched">Mismatched</option>
              <option value="formatting_error">Formatting Errors</option>
              <option value="unverifiable">Unverifiable</option>
            </select>
          </div>

          <div>
            <label htmlFor="sort" className="text-sm font-medium text-gray-700 mr-2">
              Sort by:
            </label>
            <select
              id="sort"
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value)}
              className="text-sm border-gray-300 rounded-md"
            >
              <option value="confidence">Confidence</option>
              <option value="slide">Slide Number</option>
              <option value="status">Status</option>
            </select>
          </div>
        </div>

        <div className="text-sm text-gray-600">
          Showing {filteredDiscrepancies.length} of {discrepancies.length} items
        </div>
      </div>

      {/* Results List */}
      <div className="space-y-3">
        {filteredDiscrepancies.map((item, index) => (
          <div
            key={index}
            className={`border rounded-lg p-4 ${getStatusColor(item.validation_status)}`}
          >
            <div className="flex items-start justify-between">
              <div className="flex items-start space-x-3 flex-1">
                {getStatusIcon(item.validation_status)}
                <div className="flex-1 min-w-0">
                  <div className="flex items-center justify-between">
                    <h4 className="text-sm font-medium text-gray-900">
                      Slide {item?.pdf_value_id?.split("_")?.[1] ?? ""}
                    </h4>
                    <span className="text-xs text-gray-500">
                      Confidence: {Math.round((item.confidence || 0) * 100)}%
                    </span>
                  </div>

                  <div className="mt-2 grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-xs font-medium text-gray-700">Presentation Value</p>
                      <p className="text-sm text-gray-900">{item.pdf_value}</p>
                      <p className="text-xs text-gray-500">{item.pdf_context}</p>
                    </div>
                    <div>
                      <p className="text-xs font-medium text-gray-700">Excel Value</p>
                      <p className="text-sm text-gray-900">
                        {item?.excel_match?.excel_value}
                      </p>
                    </div>
                  </div>

                  {item.validation_status !== 'matched' && (
                    <div className="mt-2">
                      <span className="inline-flex items-center px-2 py-1 rounded text-xs font-medium bg-white bg-opacity-50">
                        {item.discrepancy_type?.replace('_', ' ').toUpperCase()}
                      </span>
                    </div>
                  )}
                </div>
              </div>

              <button
                onClick={() => toggleExpanded(index)}
                className="ml-2 p-1 text-gray-400 hover:text-gray-600"
              >
                {expandedItems.has(index) ? (
                  <ChevronUp className="h-4 w-4" />
                ) : (
                  <ChevronDown className="h-4 w-4" />
                )}
              </button>
            </div>

            {expandedItems.has(index) && (
              <div className="mt-4 pt-4 border-t border-gray-200 border-opacity-50">
                <div className="space-y-3">
                  <div>
                    <h5 className="text-xs font-medium text-gray-700 mb-1">Audit Reasoning</h5>
                    <p className="text-sm text-gray-800">{item.audit_reasoning}</p>
                  </div>
                </div>
              </div>
            )}
          </div>
        ))}
      </div>

      {filteredDiscrepancies.length === 0 && (
        <div className="text-center py-12">
          <Eye className="mx-auto h-12 w-12 text-gray-400" />
          <h3 className="mt-2 text-sm font-medium text-gray-900">No items found</h3>
          <p className="mt-1 text-sm text-gray-500">
            Try adjusting your filter or sort options.
          </p>
        </div>
      )}
    </div>
  );
};

export default DiscrepancyList;