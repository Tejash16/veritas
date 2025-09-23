import React, { useState, useEffect } from 'react';
import { useParams, useLocation, useNavigate } from 'react-router-dom';
import { Pie, Bar } from 'react-chartjs-2';
import {
  Chart as ChartJS,
  ArcElement,
  Tooltip,
  Legend,
  CategoryScale,
  LinearScale,
  BarElement,
  Title
} from 'chart.js';
import { AlertTriangle, CheckCircle, XCircle, Eye, FileText } from 'lucide-react';

import DiscrepancyList from './DiscrepancyList';
import MetricsPanel from './MetricsPanel';

ChartJS.register(
  ArcElement,
  Tooltip,
  Legend,
  CategoryScale,
  LinearScale,
  BarElement,
  Title
);

const AuditDashboard = () => {
  const { sessionId } = useParams();
  const location = useLocation();
  const navigate = useNavigate();
  
  const [auditResults, setAuditResults] = useState(null);
  const [selectedTab, setSelectedTab] = useState('overview');

  useEffect(() => {
    if (location.state?.auditResults) {
      setAuditResults(location.state.auditResults);
    }
  }, [location.state]);

  const generateReportData = () => {
    if (!auditResults) return null;

    const pieData = {
      labels: ['Matched', 'Mismatched', 'Formatting Errors', 'Unverifiable'],
      datasets: [
        {
          data: [
            auditResults.summary.matched,
            auditResults.summary.mismatched,
            auditResults.summary.formatting_errors,
            auditResults.summary.unverifiable
          ],
          backgroundColor: [
            '#10B981', // Green
            '#EF4444', // Red
            '#F59E0B', // Yellow
            '#6B7280'  // Gray
          ],
          borderWidth: 2,
          borderColor: '#ffffff'
        }
      ]
    };

    const barData = {
      labels: ['Overall Accuracy', 'Data Quality Score'],
      datasets: [
        {
          label: 'Percentage',
          data: [
            auditResults.summary.overall_accuracy,
            85 // Mock data quality score
          ],
          backgroundColor: ['#3B82F6', '#8B5CF6'],
          borderColor: ['#2563EB', '#7C3AED'],
          borderWidth: 1
        }
      ]
    };

    return { pieData, barData };
  };

  const getRiskLevelColor = (risk) => {
    switch (risk) {
      case 'low': return 'text-green-600 bg-green-100';
      case 'medium': return 'text-yellow-600 bg-yellow-100';
      case 'high': return 'text-red-600 bg-red-100';
      default: return 'text-gray-600 bg-gray-100';
    }
  };

  if (!auditResults) {
    return (
      <div className="flex items-center justify-center min-h-96">
        <div className="animate-spin rounded-full h-32 w-32 border-b-2 border-blue-600"></div>
      </div>
    );
  }

  const chartData = generateReportData();

  return (
    <div className="max-w-7xl mx-auto space-y-6">
      {/* Header */}
      <div className="bg-white shadow rounded-lg">
        <div className="px-4 py-5 sm:p-6">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold text-gray-900">Audit Results</h1>
              <p className="mt-1 text-sm text-gray-600">Session ID: {sessionId}</p>
            </div>
            <div className="flex items-center space-x-4">
              <span className={`inline-flex items-center px-3 py-1 rounded-full text-sm font-medium ${getRiskLevelColor(auditResults.risk_assessment)}`}>
                Risk Level: {auditResults.risk_assessment.toUpperCase()}
              </span>
              <button
                onClick={() => navigate(`/reports/${sessionId}`)}
                className="px-4 py-2 text-sm font-medium text-white bg-blue-600 border border-transparent rounded-md hover:bg-blue-700"
              >
                <FileText className="inline h-4 w-4 mr-2" />
                Generate Report
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
        <div className="bg-white overflow-hidden shadow rounded-lg">
          <div className="p-5">
            <div className="flex items-center">
              <div className="flex-shrink-0">
                <CheckCircle className="h-6 w-6 text-green-400" />
              </div>
              <div className="ml-5 w-0 flex-1">
                <dl>
                  <dt className="text-sm font-medium text-gray-500 truncate">
                    Matched Values
                  </dt>
                  <dd className="text-lg font-medium text-gray-900">
                    {auditResults.summary.matched}
                  </dd>
                </dl>
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white overflow-hidden shadow rounded-lg">
          <div className="p-5">
            <div className="flex items-center">
              <div className="flex-shrink-0">
                <XCircle className="h-6 w-6 text-red-400" />
              </div>
              <div className="ml-5 w-0 flex-1">
                <dl>
                  <dt className="text-sm font-medium text-gray-500 truncate">
                    Mismatched Values
                  </dt>
                  <dd className="text-lg font-medium text-gray-900">
                    {auditResults.summary.mismatched}
                  </dd>
                </dl>
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white overflow-hidden shadow rounded-lg">
          <div className="p-5">
            <div className="flex items-center">
              <div className="flex-shrink-0">
                <AlertTriangle className="h-6 w-6 text-yellow-400" />
              </div>
              <div className="ml-5 w-0 flex-1">
                <dl>
                  <dt className="text-sm font-medium text-gray-500 truncate">
                    Formatting Errors
                  </dt>
                  <dd className="text-lg font-medium text-gray-900">
                    {auditResults.summary.formatting_errors}
                  </dd>
                </dl>
              </div>
            </div>
          </div>
        </div>

        <div className="bg-white overflow-hidden shadow rounded-lg">
          <div className="p-5">
            <div className="flex items-center">
              <div className="flex-shrink-0">
                <div className="h-6 w-6 bg-blue-500 rounded-full flex items-center justify-center text-white text-xs font-bold">
                  %
                </div>
              </div>
              <div className="ml-5 w-0 flex-1">
                <dl>
                  <dt className="text-sm font-medium text-gray-500 truncate">
                    Overall Accuracy
                  </dt>
                  <dd className="text-lg font-medium text-gray-900">
                    {Math.round(auditResults.summary.overall_accuracy)}%
                  </dd>
                </dl>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Tab Navigation */}
      <div className="bg-white shadow rounded-lg">
        <div className="border-b border-gray-200">
          <nav className="-mb-px flex space-x-8 px-6">
            {[
              { id: 'overview', name: 'Overview' },
              { id: 'discrepancies', name: 'Discrepancies' },
              { id: 'metrics', name: 'Metrics' }
            ].map((tab) => (
              <button
                key={tab.id}
                onClick={() => setSelectedTab(tab.id)}
                className={`py-4 px-1 border-b-2 font-medium text-sm ${
                  selectedTab === tab.id
                    ? 'border-blue-500 text-blue-600'
                    : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                }`}
              >
                {tab.name}
              </button>
            ))}
          </nav>
        </div>

        <div className="p-6">
          {selectedTab === 'overview' && (
            <div className="space-y-6">
              {/* Charts */}
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <div className="bg-gray-50 p-4 rounded-lg">
                  <h3 className="text-lg font-medium text-gray-900 mb-4">Validation Results</h3>
                  <div className="h-64">
                    {chartData?.pieData && <Pie
                      data={chartData.pieData}
                      options={{
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                          legend: {
                            position: 'bottom',
                          },
                        },
                      }}
                    />}
                  </div>
                </div>

                <div className="bg-gray-50 p-4 rounded-lg">
                  <h3 className="text-lg font-medium text-gray-900 mb-4">Quality Metrics</h3>
                  <div className="h-64">
                    {chartData?.barData && <Bar
                      data={chartData.barData}
                      options={{
                        responsive: true,
                        maintainAspectRatio: false,
                        scales: {
                          y: {
                            beginAtZero: true,
                            max: 100,
                          },
                        },
                      }}
                    />}
                  </div>
                </div>
              </div>

              {/* Recommendations */}
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
                <h3 className="text-lg font-medium text-blue-900 mb-2">Recommendations</h3>
                <ul className="space-y-1">
                  {/* Add a check here for auditResults.recommendations */}
                  {auditResults.recommendations && auditResults.recommendations.map((recommendation, index) => (
                    <li key={index} className="text-sm text-blue-800">
                      • {recommendation}
                    </li>
                  ))}
                </ul>
              </div>
            </div>
          )}

          {selectedTab === 'discrepancies' && (
            // Add a check here for auditResults.detailed_results
            <DiscrepancyList discrepancies={auditResults.detailed_results || []} />
          )}

          {selectedTab === 'metrics' && (
            <MetricsPanel sessionId={sessionId} />
          )}
        </div>
      </div>
    </div>
  );
};

export default AuditDashboard;