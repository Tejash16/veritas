import React, { useState, useEffect } from 'react';
import { useParams, useLocation, useNavigate } from 'react-router-dom';
import { Pie } from 'react-chartjs-2';
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
import Preview from './Preview';
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
      console.log('[Audit Dashboard] Received audit results:', auditResults);
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

    return { pieData };
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
          <div className="flex flex-col items-center space-y-4">
            <div>
              <h1 className="text-2xl font-bold text-gray-900 text-center">Audit Results</h1>
            </div>
            <div className="flex space-x-4 w-full">
              {/* Summary Cards in Header */}
              <div className="bg-white overflow-hidden shadow rounded-lg flex-1">
                <div className="p-5">
                  <div className="flex flex-col items-center text-center">
                    <div className="flex-shrink-0">
                      <CheckCircle className="h-6 w-6 text-green-400" />
                    </div>
                    <div className="mt-3">
                      <dl>
                        <dt className="text-sm font-medium text-gray-500">
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

              <div className="bg-white overflow-hidden shadow rounded-lg flex-1">
                <div className="p-5">
                  <div className="flex flex-col items-center text-center">
                    <div className="flex-shrink-0">
                      <XCircle className="h-6 w-6 text-red-400" />
                    </div>
                    <div className="mt-3">
                      <dl>
                        <dt className="text-sm font-medium text-gray-500">
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

              <div className="bg-white overflow-hidden shadow rounded-lg flex-1">
                <div className="p-5">
                  <div className="flex flex-col items-center text-center">
                    <div className="flex-shrink-0">
                      <AlertTriangle className="h-6 w-6 text-gray-600" />
                    </div>
                    <div className="mt-3">
                      <dl>
                        <dt className="text-sm font-medium text-gray-500">
                          Unverified Values
                        </dt>
                        <dd className="text-lg font-medium text-gray-900">
                          {auditResults.summary.unverifiable}
                        </dd>
                      </dl>
                    </div>
                  </div>
                </div>
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
              { id: 'preview', name: 'Preview' }
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
              <div className="grid grid-cols-1 gap-6">
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
              </div>
            </div>
          )}

          {selectedTab === 'discrepancies' && (
            <DiscrepancyList discrepancies={auditResults.detailed_results || []} />
          )}

          {selectedTab === 'preview' && (
            <Preview preview={auditResults.detailed_results} sessionId={sessionId} />
          )}
        </div>
      </div>
    </div>
  );
};

export default AuditDashboard;