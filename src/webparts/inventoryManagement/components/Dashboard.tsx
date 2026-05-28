import * as React from 'react';
import styles from './Dashboard.module.scss';
import { IInventoryItem } from '../models/IInventoryItem';
import { IRequest } from '../models/IRequest';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';

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

  // --- Utility Helpers ---
  const getFluentColor = (status: string, alpha: number = 1): string => {
    const s = status.toLowerCase();
    if (s.includes('in stock') || s === 'yes' || s === 'approved' || s === 'available') {
      return `rgba(16, 124, 16, ${alpha})`; // Fluent Green
    }
    if (s.includes('assigned') || s === 'active' || s === 'in use' || s === 'assigned to employee') {
      return `rgba(0, 120, 212, ${alpha})`; // Fluent Blue
    }
    if (s.includes('pending') || s.includes('awaiting')) {
      return `rgba(255, 185, 0, ${alpha})`; // Fluent Gold/Yellow
    }
    if (s.includes('damaged') || s.includes('rejected') || s.includes('lost') || s.includes('broken')) {
      return `rgba(216, 59, 1, ${alpha})`; // Fluent Red
    }
    if (s.includes('borrowed') || s.includes('requested') || s === 'requested asset') {
      return `rgba(135, 100, 184, ${alpha})`; // Fluent Purple
    }
    
    const hash = status.split('').reduce((acc, char) => acc + char.charCodeAt(0), 0);
    const colors = [
      `rgba(0, 120, 212, ${alpha})`, // Blue
      `rgba(16, 124, 16, ${alpha})`,  // Green
      `rgba(135, 100, 184, ${alpha})`, // Purple
      `rgba(0, 130, 114, ${alpha})`,  // Teal (#008272)
      `rgba(216, 59, 1, ${alpha})`,   // Orange
    ];
    return colors[hash % colors.length];
  };

  const formatDate = (dateStr: string): string => {
    if (!dateStr) return 'N/A';
    try {
      return dateStr.split('T')[0];
    } catch {
      return dateStr;
    }
  };

  const renderStatusPill = (status: string): JSX.Element => {
    const s = (status || '').toLowerCase();
    if (s.includes('approved') || s === 'active' || s === 'assigned' || s === 'in stock' || s === 'yes') {
      return <span className={`${styles.statusBadge} ${styles.badgeApproved}`}>{status}</span>;
    }
    if (s.includes('declined') || s.includes('rejected') || s === 'no' || s.includes('damaged')) {
      return <span className={`${styles.statusBadge} ${styles.badgeDeclined}`}>{status}</span>;
    }
    return <span className={`${styles.statusBadge} ${styles.badgePending}`}>{status || 'Pending'}</span>;
  };

  // --- Advanced Telemetry Calculations ---
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

  // 1. Asset Allocation Rate (Admin)
  const assignedAssetsCount = totalAssets - availableAssets;
  const allocationRate = totalAssets > 0 ? ((assignedAssetsCount / totalAssets) * 100).toFixed(0) : '0';
  const stockPercentage = totalAssets > 0 ? ((availableAssets / totalAssets) * 100).toFixed(0) : '0';

  // 2. Request Fulfillment Success Rate (Manager & Employee)
  const decidedRequests = requests.filter(r => r.status === 'Approved' || r.status === 'Declined');
  const approvedCount = requests.filter(r => r.status === 'Approved').length;
  const approvalSuccessRate = decidedRequests.length > 0
    ? ((approvedCount / decidedRequests.length) * 100).toFixed(0)
    : '100';

  // --- Data Processing for Charts ---

  // 1. Primary status counts
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
  const primaryPieTitle = isManagerView ? 'Approvals Queue Status' : 'Asset Status Distribution';
  const primaryPieSubtitle = isManagerView ? 'Requests categorized by manager approval state' : 'Current condition and status of registered assets';

  const statusLabels = Object.keys(statusCounts).length ? Object.keys(statusCounts) : ['No data'];
  const statusDataValues = Object.keys(statusCounts).length
    ? Object.keys(statusCounts).map(k => statusCounts[k])
    : [1];

  const assetStatusData = {
    labels: statusLabels,
    datasets: [
      {
        label: primaryPieLabel,
        data: statusDataValues,
        backgroundColor: statusLabels.map(label => getFluentColor(label, 0.75)),
        borderColor: statusLabels.map(label => getFluentColor(label, 1.0)),
        borderWidth: 1.5,
      },
    ],
  };

  // 2. Assets by Type (Bar Chart)
  const typeCounts = items.reduce((acc, item) => {
    const type = item.assetType || 'Unknown';
    acc[type] = (acc[type] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  const assetTypeLabels = Object.keys(typeCounts);
  const assetTypeDataValues = Object.keys(typeCounts).map(k => typeCounts[k]);

  const assetTypeData = {
    labels: assetTypeLabels.length ? assetTypeLabels : ['No assets'],
    datasets: [
      {
        label: 'Number of Assets',
        data: assetTypeDataValues.length ? assetTypeDataValues : [0],
        backgroundColor: 'rgba(0, 120, 212, 0.75)',
        borderColor: 'rgba(0, 120, 212, 1)',
        borderWidth: 1.5,
        hoverBackgroundColor: 'rgba(0, 90, 158, 0.85)',
        hoverBorderColor: 'rgba(0, 90, 158, 1)',
        borderRadius: 4,
      },
    ],
  };

  // 3. Doughnut request status counts
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
      ? ['No approved requests yet']
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
        backgroundColor: doughnutLabels.map(label => getFluentColor(label, 0.75)),
        borderColor: doughnutLabels.map(label => getFluentColor(label, 1.0)),
        borderWidth: 1.5,
      },
    ],
  };

  // --- Clean Segoe UI Options for Chart.js ---
  const chartPlugins = {
    legend: {
      position: 'bottom' as const,
      labels: {
        boxWidth: 10,
        boxHeight: 10,
        padding: 16,
        usePointStyle: true,
        font: {
          family: "'Segoe UI', -apple-system, sans-serif",
          size: 11,
          weight: 'normal' as const,
        },
        color: '#323130',
      },
    },
    tooltip: {
      backgroundColor: '#ffffff',
      titleColor: '#323130',
      bodyColor: '#323130',
      borderColor: 'rgba(0,0,0,0.08)',
      borderWidth: 1,
      padding: 10,
      boxPadding: 6,
      cornerRadius: 6,
      usePointStyle: true,
      titleFont: {
        family: "'Segoe UI', -apple-system, sans-serif",
        size: 12,
        weight: 'bold' as const,
      },
      bodyFont: {
        family: "'Segoe UI', -apple-system, sans-serif",
        size: 12,
      },
    },
  };

  const pieOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: chartPlugins,
  };

  const doughnutOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: chartPlugins,
    cutout: '65%',
  };

  const assetTypeOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: false,
      },
      tooltip: chartPlugins.tooltip,
    },
    scales: {
      x: {
        grid: {
          display: false,
        },
        ticks: {
          font: {
            family: "'Segoe UI', -apple-system, sans-serif",
            size: 11,
          },
          color: '#605e5c',
        },
      },
      y: {
        grid: {
          color: 'rgba(0,0,0,0.04)',
        },
        ticks: {
          precision: 0,
          font: {
            family: "'Segoe UI', -apple-system, sans-serif",
            size: 11,
          },
          color: '#605e5c',
        },
      },
    },
  };

  // --- Action Center Data Filtering ---
  // Admins: Manager approved requests awaiting physical assignment
  const pendingAssignments = requests.filter(r => r.status === 'Approved' && r.assetStatus === 'Pending');

  // Managers: Unresolved requests awaiting manager approval
  const pendingApprovals = requests.filter(r => (r.status || '').toLowerCase() === 'pending');

  // Employees: Active personal requests
  const activePersonalRequests = requests.filter(r => r.status === 'Pending' || (r.status === 'Approved' && r.assetStatus === 'Pending'));

  return (
    <div className={styles.dashboard}>
      {/* SharePoint Status Banners */}
      {isManagerView && (
        <div className={styles.dashboardIntro}>
          <MessageBar messageBarType={MessageBarType.info}>
            <strong>Inventory Manager Dashboard</strong> — Visual metrics and approval queues are aggregated from request lists. Full data resides in the <strong>Approvals</strong> registry.
          </MessageBar>
        </div>
      )}
      {isAdmin && (
        <div className={styles.dashboardIntro}>
          <MessageBar messageBarType={MessageBarType.success}>
            <strong>Administrator Dashboard</strong> — Analytics are derived directly from the physical inventory items. Assignment metrics display admin-approved items.
          </MessageBar>
        </div>
      )}
      {!isAdmin && !isInventoryManager && (
        <div className={styles.dashboardIntro}>
          <MessageBar messageBarType={MessageBarType.info}>
            <strong>Personal Asset Hub</strong> — Real-time telemetry tracking your assigned devices and ongoing request status.
          </MessageBar>
        </div>
      )}

      {/* Modern KPI Summary Grid */}
      <div className={styles.summaryGrid}>
        {/* Card 1: Total Assets */}
        <div className={`${styles.summaryCard} ${styles.cardBlue}`}>
          <div className={styles.iconContainer}>
            <Icon iconName="Package" />
          </div>
          <div className={styles.cardInfo}>
            <span className={styles.summaryValue}>{totalAssets}</span>
            <span className={styles.summaryLabel}>Total Assets</span>
            <span className={styles.summarySubtitle}>
              {isAdmin ? `${totalAssets} items • ${allocationRate}% allocated` : `${totalAssets} item${totalAssets === 1 ? '' : 's'} in registry`}
            </span>
          </div>
        </div>

        {/* Card 2: Available Assets */}
        <div className={`${styles.summaryCard} ${styles.cardGreen}`}>
          <div className={styles.iconContainer}>
            <Icon iconName="Accept" />
          </div>
          <div className={styles.cardInfo}>
            <span className={styles.summaryValue}>{availableAssets}</span>
            <span className={styles.summaryLabel}>Available Assets</span>
            <span className={styles.summarySubtitle}>
              {availableAssets} in stock ({stockPercentage}% of total)
            </span>
          </div>
        </div>

        {/* Card 3: Requests in queue / Total Requests */}
        <div className={`${styles.summaryCard} ${styles.cardPurple}`}>
          <div className={styles.iconContainer}>
            <Icon iconName="Send" />
          </div>
          <div className={styles.cardInfo}>
            <span className={styles.summaryValue}>{totalRequests}</span>
            <span className={styles.summaryLabel}>
              {isManagerView ? 'Requests in Queue' : 'Total Requests'}
            </span>
            <span className={styles.summarySubtitle}>
              {!isAdmin && !isInventoryManager 
                ? `${totalRequests} requested • ${approvalSuccessRate}% approval`
                : `${totalRequests} submissions • ${approvalSuccessRate}% approval`}
            </span>
          </div>
        </div>

        {/* Card 4: Awaiting Approval / Pending Requests */}
        <div className={`${styles.summaryCard} ${styles.cardGold}`}>
          <div className={styles.iconContainer}>
            <Icon iconName="Clock" />
          </div>
          <div className={styles.cardInfo}>
            <span className={styles.summaryValue}>
              {isManagerView ? awaitingManagerDecision : pendingRequests}
            </span>
            <span className={styles.summaryLabel}>
              {isManagerView ? 'Awaiting Approval' : 'Pending Requests'}
            </span>
            <span className={styles.summarySubtitle}>
              {isManagerView
                ? `${awaitingManagerDecision} requires review`
                : `${pendingRequests} awaiting processing`}
            </span>
          </div>
        </div>
      </div>

      {/* Modern Dashboard Charts Grid */}
      <div className={styles.chartsGrid}>
        {/* Chart 1: Primary Status */}
        <div className={styles.chartCard}>
          <div className={styles.chartHeader}>
            <h3>{primaryPieTitle}</h3>
            <span className={styles.chartSubtitle}>{primaryPieSubtitle}</span>
          </div>
          <div className={styles.chartContainer}>
            <Pie data={assetStatusData} options={pieOptions} />
          </div>
        </div>

        {/* Chart 2: Types */}
        <div className={styles.chartCard}>
          <div className={styles.chartHeader}>
            <h3>Assets by Type</h3>
            <span className={styles.chartSubtitle}>Categorized distribution of equipment</span>
          </div>
          <div className={styles.chartContainer}>
            <Bar data={assetTypeData} options={assetTypeOptions} />
          </div>
        </div>

        {/* Chart 3: Doughnut (Request Status) */}
        <div className={styles.chartCard}>
          <div className={styles.chartHeader}>
            <h3>
              {isManagerView ? 'Post-Approval Assignment Status' : 'Request Fulfillment Status'}
            </h3>
            <span className={styles.chartSubtitle}>
              {isManagerView
                ? 'Status of asset handouts for manager-approved requests'
                : 'Current status across all request pipelines'}
            </span>
          </div>
          <div className={styles.chartContainer}>
            <Doughnut data={requestStatusData} options={doughnutOptions} />
          </div>
        </div>
      </div>

      {/* --- Action Centers (Bottom panels tailored per role) --- */}

      {/* A. ADMIN ACTION CENTER: Approved Requests Awaiting IT Asset Assignment */}
      {isAdmin && (
        <div className={styles.actionCard}>
          <div className={styles.actionHeader}>
            <div className={styles.headerText}>
              <h3>IT Fulfillment Action Center</h3>
              <span className={styles.actionSubtitle}>Manager-approved requests requiring physical device assignment and handout</span>
            </div>
            <span className={styles.headerBadge}>{pendingAssignments.length} Awaiting Handout</span>
          </div>
          <div className={styles.tableWrapper}>
            {pendingAssignments.length === 0 ? (
              <div className={styles.emptyState}>
                <Icon iconName="Completed" />
                <span>All approved requests have been physically assigned! Excellent job.</span>
              </div>
            ) : (
              <table className={styles.actionTable}>
                <thead>
                  <tr>
                    <th>Requester</th>
                    <th>Equipment Type</th>
                    <th>Qty</th>
                    <th>Date Approved</th>
                    <th>Fulfillment Status</th>
                  </tr>
                </thead>
                <tbody>
                  {pendingAssignments.slice(0, 5).map(req => (
                    <tr key={req.id}>
                      <td><strong>{req.requesterName}</strong></td>
                      <td>{req.assetTitle}</td>
                      <td>{req.quantity}</td>
                      <td>{formatDate(req.requestDate)}</td>
                      <td>{renderStatusPill('Awaiting Handout')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* B. MANAGER ACTION CENTER: Requests Awaiting Manager Review */}
      {isManagerView && (
        <div className={styles.actionCard}>
          <div className={styles.actionHeader}>
            <div className={styles.headerText}>
              <h3>Manager Approval Queue</h3>
              <span className={styles.actionSubtitle}>Incoming equipment requests awaiting your review and approval</span>
            </div>
            <span className={styles.headerBadge}>{pendingApprovals.length} Pending Review</span>
          </div>
          <div className={styles.tableWrapper}>
            {pendingApprovals.length === 0 ? (
              <div className={styles.emptyState}>
                <Icon iconName="Completed" />
                <span>Excellent! You have cleared all pending requests.</span>
              </div>
            ) : (
              <table className={styles.actionTable}>
                <thead>
                  <tr>
                    <th>Requester</th>
                    <th>Requested Item</th>
                    <th>Qty</th>
                    <th>Date Submitted</th>
                    <th>Purpose / Business Reason</th>
                    <th>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {pendingApprovals.slice(0, 5).map(req => (
                    <tr key={req.id}>
                      <td><strong>{req.requesterName}</strong></td>
                      <td>{req.assetTitle}</td>
                      <td>{req.quantity}</td>
                      <td>{formatDate(req.requestDate)}</td>
                      <td><em>{req.reason || 'No justification provided'}</em></td>
                      <td>{renderStatusPill('Pending Review')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        </div>
      )}

      {/* C. EMPLOYEE ACTION CENTER: Personal Requests Tracker & Assigned Devices */}
      {!isAdmin && !isInventoryManager && (
        <React.Fragment>
          {/* Employee active requests */}
          <div className={styles.actionCard}>
            <div className={styles.actionHeader}>
              <div className={styles.headerText}>
                <h3>My Active Requests Tracker</h3>
                <span className={styles.actionSubtitle}>Real-time tracking of your pending asset requisitions</span>
              </div>
              <span className={styles.headerBadge}>{activePersonalRequests.length} Active</span>
            </div>
            <div className={styles.tableWrapper}>
              {activePersonalRequests.length === 0 ? (
                <div className={styles.emptyState}>
                  <Icon iconName="Info" />
                  <span>You have no active equipment requests at this time.</span>
                </div>
              ) : (
                <table className={styles.actionTable}>
                  <thead>
                    <tr>
                      <th>Requested Equipment</th>
                      <th>Qty</th>
                      <th>Date Requested</th>
                      <th>Manager Decision</th>
                      <th>IT Fulfillment</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activePersonalRequests.map(req => (
                      <tr key={req.id}>
                        <td><strong>{req.assetTitle}</strong></td>
                        <td>{req.quantity}</td>
                        <td>{formatDate(req.requestDate)}</td>
                        <td>{renderStatusPill(req.status)}</td>
                        <td>
                          {req.status === 'Approved' 
                            ? renderStatusPill('Awaiting Handout') 
                            : renderStatusPill('Awaiting Manager Approval')}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>

          {/* Employee assigned devices */}
          <div className={styles.actionCard}>
            <div className={styles.actionHeader}>
              <div className={styles.headerText}>
                <h3>My Assigned Devices</h3>
                <span className={styles.actionSubtitle}>Equipment currently assigned and in your possession</span>
              </div>
              <span className={styles.headerBadge}>{items.length} Assigned</span>
            </div>
            <div className={styles.tableWrapper}>
              {items.length === 0 ? (
                <div className={styles.emptyState}>
                  <Icon iconName="DeviceOff" />
                  <span>You do not have any registered devices assigned to you.</span>
                </div>
              ) : (
                <table className={styles.actionTable}>
                  <thead>
                    <tr>
                      <th>Equipment Name</th>
                      <th>Category</th>
                      <th>Serial Number</th>
                      <th>Status State</th>
                      <th>Notes / Specification</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map(item => (
                      <tr key={item.id}>
                        <td><strong>{item.title || item.assetName}</strong></td>
                        <td>{item.assetType}</td>
                        <td><code>{item.serialNumber || 'N/A'}</code></td>
                        <td>{renderStatusPill(item.status || 'Active')}</td>
                        <td><em>{item.note || 'No additional details'}</em></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </React.Fragment>
      )}
    </div>
  );
};
