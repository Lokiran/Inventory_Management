import * as React from 'react';
import styles from './Dashboard.module.scss';
import { IInventoryItem } from '../models/IInventoryItem';
import { IRequest } from '../models/IRequest';

import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ArcElement,
} from 'chart.js';
import { Bar, Pie, Doughnut } from 'react-chartjs-2';

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ArcElement
);

export interface IDashboardProps {
  items: IInventoryItem[];
  requests: IRequest[];
  isAdmin?: boolean;
  /** When true, dashboard copy and the primary pie chart follow the Approvals queue (requests), not inventory asset status. */
  isInventoryManager?: boolean;
}

export const Dashboard: React.FunctionComponent<IDashboardProps> = (props) => {
  const { items, requests, isAdmin, isInventoryManager } = props;
  const isManagerView = !!isInventoryManager && !isAdmin;

  // --- Data Processing for Charts ---

  // 1. Primary pie: Admin / Employee = asset status from inventory; Inventory Manager = request status from Approvals queue (same fields as Approvals tab)
  const statusCounts = isManagerView
    ? requests.reduce((acc, req) => {
        const status = req.status || 'Pending';
        acc[status] = (acc[status] || 0) + 1;
        return acc;
      }, {} as Record<string, number>)
    : items.reduce((acc, item) => {
        const status = item.status || 'Unknown';
        acc[status] = (acc[status] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

  const primaryPieLabel = isManagerView ? 'Requests in approval queue' : 'Assets by Status';
  const primaryPieTitle = isManagerView ? 'Approvals queue by request status' : 'Asset Status Distribution';

  const assetStatusData = {
    labels: Object.keys(statusCounts).length ? Object.keys(statusCounts) : ['No data'],
    datasets: [
      {
        label: primaryPieLabel,
        data: Object.keys(statusCounts).length
          ? Object.keys(statusCounts).map(k => statusCounts[k])
          : [1],
        backgroundColor: [
          'rgba(54, 162, 235, 0.6)',
          'rgba(255, 206, 86, 0.6)',
          'rgba(75, 192, 192, 0.6)',
          'rgba(255, 99, 132, 0.6)',
          'rgba(153, 102, 255, 0.6)',
          'rgba(255, 159, 64, 0.6)',
        ],
        borderColor: [
          'rgba(54, 162, 235, 1)',
          'rgba(255, 206, 86, 1)',
          'rgba(75, 192, 192, 1)',
          'rgba(255, 99, 132, 1)',
          'rgba(153, 102, 255, 1)',
          'rgba(255, 159, 64, 1)',
        ],
        borderWidth: 1,
      },
    ],
  };

  // 2. Assets by Type (Bar Chart)
  const typeCounts = items.reduce((acc, item) => {
    const type = item.assetType || 'Unknown';
    acc[type] = (acc[type] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const assetTypeData = {
    labels: Object.keys(typeCounts),
    datasets: [
      {
        label: 'Number of Assets',
        data: Object.keys(typeCounts).map(k => typeCounts[k]),
        backgroundColor: 'rgba(54, 162, 235, 0.6)',
        borderColor: 'rgba(54, 162, 235, 1)',
        borderWidth: 1,
      },
    ],
  };

  const assetTypeOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: false,
      },
      title: {
        display: false,
      },
    },
  };

  // 3. Doughnut: Admin = asset assignment status; Manager = fulfillment after manager approval (assetStatus); Employee = request status
  const requestStatusCounts = isManagerView
    ? requests
        .filter(req => (req.status || '').toLowerCase() === 'approved')
        .reduce((acc, req) => {
          const status = req.assetStatus || 'Pending';
          acc[status] = (acc[status] || 0) + 1;
          return acc;
        }, {} as Record<string, number>)
    : requests.reduce((acc, req) => {
        const status = isAdmin ? (req.assetStatus || 'Pending') : (req.status || 'Pending');
        acc[status] = (acc[status] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

  const doughnutLabels = Object.keys(requestStatusCounts).length
    ? Object.keys(requestStatusCounts)
    : isManagerView
      ? ['No manager-approved requests yet']
      : ['No data'];
  const doughnutDataValues = Object.keys(requestStatusCounts).length
    ? Object.keys(requestStatusCounts).map(k => requestStatusCounts[k])
    : [1];

  const requestStatusData = {
    labels: doughnutLabels,
    datasets: [
      {
        label: isManagerView ? 'Assignment status (approved requests)' : 'Requests by Status',
        data: doughnutDataValues,
        backgroundColor: [
          'rgba(255, 206, 86, 0.6)', // Pending
          'rgba(75, 192, 192, 0.6)', // Approved
          'rgba(255, 99, 132, 0.6)', // Rejected
          'rgba(153, 102, 255, 0.6)',
        ],
        borderColor: [
          'rgba(255, 206, 86, 1)',
          'rgba(75, 192, 192, 1)',
          'rgba(255, 99, 132, 1)',
          'rgba(153, 102, 255, 1)',
        ],
        borderWidth: 1,
      },
    ],
  };

  // --- Quick Summaries ---
  const totalAssets = items.length;
  const totalRequests = requests.length;
  const pendingRequests = requests.filter(r => {
    const status = isAdmin ? (r.assetStatus || 'Pending') : (r.status || 'Pending');
    return status === 'Pending';
  }).length;
  const availableAssets = items.filter(i => i.status === 'In Stock' || i.status === 'Yes').length;
  const awaitingManagerDecision = isManagerView
    ? requests.filter(r => (r.status || '').toLowerCase() === 'pending').length
    : 0;

  return (
    <div className={styles.dashboard}>
      {isManagerView && (
        <p className={`${styles.dashboardIntro} ${styles.dashboardIntroManager}`}>
          <strong>Inventory Manager dashboard</strong> — KPIs and the first chart use the same request list as the{' '}
          <strong>Approvals</strong> tab (full queue, not search-filtered). Administrators see a separate dashboard focused on inventory and assignment.
        </p>
      )}
      {isAdmin && (
        <p className={`${styles.dashboardIntro} ${styles.dashboardIntroAdmin}`}>
          <strong>Administrator dashboard</strong> — Asset analytics come from inventory records; request metrics reflect the assignment queue (admin-approved requests).
        </p>
      )}
      <div className={styles.summaryGrid}>
        <div className={styles.summaryCard}>
          <div className={styles.summaryValue}>{totalAssets}</div>
          <div className={styles.summaryLabel}>Total Assets</div>
        </div>
        <div className={styles.summaryCard}>
          <div className={styles.summaryValue}>{availableAssets}</div>
          <div className={styles.summaryLabel}>Available Assets</div>
        </div>
        <div className={styles.summaryCard}>
          <div className={styles.summaryValue}>{totalRequests}</div>
          <div className={styles.summaryLabel}>{isManagerView ? 'Requests in queue' : 'Total Requests'}</div>
        </div>
        <div className={styles.summaryCard}>
          <div className={styles.summaryValue}>{isManagerView ? awaitingManagerDecision : pendingRequests}</div>
          <div className={styles.summaryLabel}>
            {isManagerView ? 'Awaiting manager approval' : 'Pending Requests'}
          </div>
        </div>
      </div>

      <div className={styles.chartsGrid}>
        <div className={styles.chartCard}>
          <h3>{primaryPieTitle}</h3>
          <div className={styles.chartContainer}>
            <Pie data={assetStatusData} options={{ maintainAspectRatio: false }} />
          </div>
        </div>
        <div className={styles.chartCard}>
          <h3>Assets by Type</h3>
          <div className={styles.chartContainer}>
            <Bar data={assetTypeData} options={assetTypeOptions} />
          </div>
        </div>
        <div className={styles.chartCard}>
          <h3>{isManagerView ? 'After approval: asset assignment status' : 'Request Status'}</h3>
          <div className={styles.chartContainer}>
            <Doughnut data={requestStatusData} options={{ maintainAspectRatio: false }} />
          </div>
        </div>
      </div>
    </div>
  );
};
