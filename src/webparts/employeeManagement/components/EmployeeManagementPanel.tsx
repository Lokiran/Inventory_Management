import * as React from 'react';
import { useState } from 'react';
import {
  Stack,
  Text,
  Pivot,
  PivotItem,
  PrimaryButton,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './EmployeeManagementPanel.module.scss';
import { IEmployeeManagementProps } from './IEmployeeManagementProps';
import { Dashboard } from './Dashboard/Dashboard';
import { AssetRequestModule } from './AssetRequest/AssetRequestModule';
import { IncidentRequestModule } from './IncidentRequest/IncidentRequestModule';
import { MyRequests } from './MyRequests/MyRequests';
import { MyAssignedAssets } from './MyAssignedAssets/MyAssignedAssets';
import { IncidentHistory } from './IncidentHistory/IncidentHistory';

export const EmployeeManagementPanel: React.FC<IEmployeeManagementProps> = (
  props: IEmployeeManagementProps
) => {
  const [selectedKey, setSelectedKey] = useState<string>('home');
  const [isLoading, setIsLoading] = useState<boolean>(false);

  return (
    <div className={styles.employeePanel}>
      <div className={styles.mainContent}>
        
        <div className={styles.heroSection}>
          <div className={styles.heroText}>
            <h2>Employee Management</h2>
            <p>Welcome back, {props.userName}!</p>
            <span className={styles.smallText}>
              Manage your requests and incidents from here.
            </span>
          </div>
        </div>

        <div className={styles.actionGrid}>
          <div className={styles.actionButtonContainer}>
            <PrimaryButton 
              text="Request Asset" 
              onClick={() => setSelectedKey('request-asset')} 
              iconProps={{ iconName: 'ShoppingCart' }}
            />
          </div>
          <div className={styles.actionButtonContainer}>
            <PrimaryButton 
              text="Raise Incident" 
              onClick={() => setSelectedKey('raise-incident')} 
              iconProps={{ iconName: 'AlertSolid' }}
            />
          </div>
        </div>

        <div className={styles.card}>
          <Pivot 
            aria-label="Employee Management Views" 
            selectedKey={selectedKey} 
            onLinkClick={(item) => {
              if (item && item.props.itemKey) {
                setSelectedKey(item.props.itemKey);
              }
            }}
          >
            <PivotItem headerText="Dashboard" itemIcon="BarChart4" itemKey="home">
              <div style={{ marginTop: '20px' }}>
                <Dashboard {...props} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>
            
            <PivotItem headerText="Request Asset" itemIcon="ShoppingCart" itemKey="request-asset">
              <div style={{ marginTop: '20px' }}>
                <AssetRequestModule {...props} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="Raise Incident" itemIcon="AlertSolid" itemKey="raise-incident">
              <div style={{ marginTop: '20px' }}>
                <IncidentRequestModule {...props} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="My Requests" itemIcon="ReviewSolid" itemKey="my-requests">
              <div style={{ marginTop: '20px' }}>
                <MyRequests {...props} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="My Assigned Assets" itemIcon="CheckMark" itemKey="my-assets">
              <div style={{ marginTop: '20px' }}>
                <MyAssignedAssets {...props} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="Incident History" itemIcon="History" itemKey="incident-history">
              <div style={{ marginTop: '20px' }}>
                <IncidentHistory {...props} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>
          </Pivot>

          {isLoading && (
            <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '100px', marginTop: '20px' }}>
              <Spinner size={SpinnerSize.large} label="Loading..." />
            </Stack>
          )}
        </div>

      </div>
    </div>
  );
};
