import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  Persona,
  PersonaSize,
  Icon,
} from '@fluentui/react';
import { Card, ICardTokens } from '@uifabric/react-cards';
import { Chart as ChartJS, ArcElement, Tooltip, Legend, CategoryScale, LinearScale, PointElement, LineElement } from 'chart.js';
import { Pie, Line } from 'react-chartjs-2';
import styles from './Dashboard.module.scss';
import { IEmployeeManagementProps } from '../IEmployeeManagementProps';
import { InventoryService } from '../../services/InventoryService';

ChartJS.register(ArcElement, Tooltip, Legend, CategoryScale, LinearScale, PointElement, LineElement);

interface IDashboardStats {
  totalRequests: number;
  pendingRequests: number;
  approvedRequests: number;
  resolvedIncidents: number;
  openIncidents: number;
}

const cardTokens: ICardTokens = { childrenMargin: 12 };

export const Dashboard: React.FC<IEmployeeManagementProps & { setIsLoading: (loading: boolean) => void }> = (props) => {
  const [stats, setStats] = useState<IDashboardStats>({
    totalRequests: 0,
    pendingRequests: 0,
    approvedRequests: 0,
    resolvedIncidents: 0,
    openIncidents: 0,
  });

  useEffect(() => {
    loadDashboardStats();
  }, []);

  const loadDashboardStats = async () => {
    try {
      props.setIsLoading(true);
      const service = new InventoryService(props.apiBaseUrl);
      const dashboardData = await service.getDashboardStats(props.userEmail);
      setStats(dashboardData);
    } catch (error) {
      console.error('Error loading dashboard stats:', error);
    } finally {
      props.setIsLoading(false);
    }
  };

  const quickActionCards = [
    {
      title: 'Request Asset',
      icon: 'ShoppingCart',
      color: '#0078d4',
      description: 'Submit a new asset request',
    },
    {
      title: 'Raise Incident',
      icon: 'AlertSolid',
      color: '#e74c3c',
      description: 'Report an issue or damage',
    },
    {
      title: 'My Requests',
      icon: 'ReviewSolid',
      color: '#27ae60',
      description: `${stats.totalRequests} active requests`,
    },
    {
      title: 'My Assets',
      icon: 'CheckMark',
      color: '#f39c12',
      description: 'View assigned assets',
    },
  ];

  const pieChartData = {
    labels: ['Pending', 'Approved', 'Rejected'],
    datasets: [
      {
        label: 'Request Status',
        data: [stats.pendingRequests, stats.approvedRequests, stats.totalRequests - stats.pendingRequests - stats.approvedRequests],
        backgroundColor: ['#ffb81c', '#107c10', '#d13438'],
        borderColor: ['#fff', '#fff', '#fff'],
        borderWidth: 2,
      },
    ],
  };

  const lineChartData = {
    labels: ['Week 1', 'Week 2', 'Week 3', 'Week 4'],
    datasets: [
      {
        label: 'Resolved Incidents',
        data: [2, 5, 3, 8],
        borderColor: '#0078d4',
        backgroundColor: 'rgba(0, 120, 212, 0.1)',
        tension: 0.4,
      },
    ],
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      {/* Welcome Section */}
      <Card>
        <Card.Section tokens={cardTokens}>
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 15 }}>
            <Persona
              imageUrl=""
              text={props.userName}
              secondaryText={props.userEmail}
              size={PersonaSize.size72}
            />
            <Stack tokens={{ childrenGap: 5 }}>
              <Text variant="xLarge" block style={{ fontWeight: 600 }}>
                Welcome, {props.userName}!
              </Text>
              <Text variant="medium" block style={{ color: '#666' }}>
                Manage your assets and incidents from here.
              </Text>
            </Stack>
          </Stack>
        </Card.Section>
      </Card>

      {/* Quick Action Cards */}
      <div>
        <Text variant="large" block style={{ fontWeight: 600, marginBottom: '15px' }}>
          Quick Actions
        </Text>
        <Stack horizontal wrap tokens={{ childrenGap: 15 }}>
          {quickActionCards.map((action, index) => (
            <Card key={index} style={{ flex: '1 1 calc(25% - 12px)', minWidth: '200px', cursor: 'pointer' }}>
              <Card.Section tokens={cardTokens}>
                <Stack tokens={{ childrenGap: 10 }}>
                  <Icon
                    iconName={action.icon}
                    style={{ fontSize: '32px', color: action.color }}
                  />
                  <Text variant="large" style={{ fontWeight: 600, color: action.color }}>
                    {action.title}
                  </Text>
                  <Text variant="small" style={{ color: '#666' }}>
                    {action.description}
                  </Text>
                </Stack>
              </Card.Section>
            </Card>
          ))}
        </Stack>
      </div>

      {/* Statistics Section */}
      <div>
        <Text variant="large" block style={{ fontWeight: 600, marginBottom: '15px' }}>
          Statistics
        </Text>
        <Stack horizontal wrap tokens={{ childrenGap: 15 }}>
          {/* Stat Cards */}
          <Card style={{ flex: '1 1 calc(25% - 12px)', minWidth: '150px' }}>
            <Card.Section tokens={cardTokens}>
              <Stack tokens={{ childrenGap: 5 }}>
                <Text variant="small" style={{ color: '#666' }}>
                  Total Requests
                </Text>
                <Text variant="xxLarge" style={{ fontWeight: 700, color: '#0078d4' }}>
                  {stats.totalRequests}
                </Text>
              </Stack>
            </Card.Section>
          </Card>

          <Card style={{ flex: '1 1 calc(25% - 12px)', minWidth: '150px' }}>
            <Card.Section tokens={cardTokens}>
              <Stack tokens={{ childrenGap: 5 }}>
                <Text variant="small" style={{ color: '#666' }}>
                  Pending
                </Text>
                <Text variant="xxLarge" style={{ fontWeight: 700, color: '#ffb81c' }}>
                  {stats.pendingRequests}
                </Text>
              </Stack>
            </Card.Section>
          </Card>

          <Card style={{ flex: '1 1 calc(25% - 12px)', minWidth: '150px' }}>
            <Card.Section tokens={cardTokens}>
              <Stack tokens={{ childrenGap: 5 }}>
                <Text variant="small" style={{ color: '#666' }}>
                  Open Incidents
                </Text>
                <Text variant="xxLarge" style={{ fontWeight: 700, color: '#e74c3c' }}>
                  {stats.openIncidents}
                </Text>
              </Stack>
            </Card.Section>
          </Card>

          <Card style={{ flex: '1 1 calc(25% - 12px)', minWidth: '150px' }}>
            <Card.Section tokens={cardTokens}>
              <Stack tokens={{ childrenGap: 5 }}>
                <Text variant="small" style={{ color: '#666' }}>
                  Resolved
                </Text>
                <Text variant="xxLarge" style={{ fontWeight: 700, color: '#27ae60' }}>
                  {stats.resolvedIncidents}
                </Text>
              </Stack>
            </Card.Section>
          </Card>
        </Stack>
      </div>

      {/* Charts */}
      <Stack horizontal tokens={{ childrenGap: 20 }}>
        <Card style={{ flex: 1 }}>
          <Card.Section tokens={cardTokens}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: '15px' }} block>
              Request Status Distribution
            </Text>
            <div style={{ maxHeight: '300px', display: 'flex', justifyContent: 'center' }}>
              <Pie data={pieChartData} options={{ maintainAspectRatio: false }} />
            </div>
          </Card.Section>
        </Card>

        <Card style={{ flex: 1 }}>
          <Card.Section tokens={cardTokens}>
            <Text variant="large" style={{ fontWeight: 600, marginBottom: '15px' }} block>
              Incident Resolution Trend
            </Text>
            <div style={{ height: '300px' }}>
              <Line data={lineChartData} options={{ maintainAspectRatio: false }} />
            </div>
          </Card.Section>
        </Card>
      </Stack>
    </Stack>
  );
};

