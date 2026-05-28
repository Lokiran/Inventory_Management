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
}

export const Dashboard: React.FunctionComponent<IDashboardProps> = (props) => {
  const { items, requests } = props;

  // --- Data Processing for Charts ---

  // 1. Asset Status Distribution (Pie Chart)
  const statusCounts = items.reduce((acc, item) => {
    const status = item.status || 'Unknown';
    acc[status] = (acc[status] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const assetStatusData = {
    labels: Object.keys(statusCounts),
    datasets: [
      {
        label: 'Assets by Status',
        data: Object.keys(statusCounts).map(k => statusCounts[k]),
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

  // 3. Request Status Overview (Doughnut Chart)
  const requestStatusCounts = requests.reduce((acc, req) => {
    const status = req.status || 'Pending';
    acc[status] = (acc[status] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const requestStatusData = {
    labels: Object.keys(requestStatusCounts),
    datasets: [
      {
        label: 'Requests by Status',
        data: Object.keys(requestStatusCounts).map(k => requestStatusCounts[k]),
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
  const pendingRequests = requests.filter(r => r.status === 'Pending').length;
  const availableAssets = items.filter(i => i.status === 'In Stock').length;

  return (
    <div className={styles.dashboard}>
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
          <div className={styles.summaryLabel}>Total Requests</div>
        </div>
        <div className={styles.summaryCard}>
          <div className={styles.summaryValue}>{pendingRequests}</div>
          <div className={styles.summaryLabel}>Pending Requests</div>
        </div>
      </div>

      <div className={styles.chartsGrid}>
        <div className={styles.chartCard}>
          <h3>Asset Status Distribution</h3>
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
          <h3>Request Status</h3>
          <div className={styles.chartContainer}>
            <Doughnut data={requestStatusData} options={{ maintainAspectRatio: false }} />
          </div>
        </div>
      </div>
    </div>
  );
};
