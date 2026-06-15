import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  Pivot,
  PivotItem,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  Persona,
  PersonaSize,
  TextField,
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
import { AdminApprovals } from './AdminApprovals/AdminApprovals';
import { InventoryService } from '../services/InventoryService';

interface IEmployeeSession {
  employeeId: string;
  employeeName: string;
  email: string;
  department: string;
}

export const EmployeeManagementPanel: React.FC<IEmployeeManagementProps> = (
  props: IEmployeeManagementProps
) => {
  const [selectedKey, setSelectedKey] = useState<string>('home');
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [employee, setEmployee] = useState<IEmployeeSession>({
    employeeId: props.userEmail || '',
    employeeName: props.userName || '',
    email: props.userEmail || '',
    department: 'General'
  });
  const [searchEmpName, setSearchEmpName] = useState<string>('');

  // Check sessionStorage for logged-in employee details
  useEffect(() => {
    const cachedSession = sessionStorage.getItem('loggedInEmployee');
    if (cachedSession) {
      try {
        const parsed = JSON.parse(cachedSession);
        setEmployee(parsed);
        setSearchEmpName(parsed.employeeName);
      } catch (e) {
        console.error('Failed to parse employee session from sessionStorage', e);
      }
    } else {
      setSearchEmpName(props.userName || '');
    }
  }, []);

  const handleEmployeeLookup = async (name: string) => {
    if (!name.trim()) return;
    try {
      setIsLoading(true);
      const service = new InventoryService(props.spContext);
      const details = await service.getEmployeeDetailsByName(name);
      
      sessionStorage.setItem('loggedInEmployee', JSON.stringify(details));
      setEmployee(details);
    } catch (e) {
      console.error('Failed to look up employee details', e);
    } finally {
      setIsLoading(false);
    }
  };

  const handleResetToSelf = () => {
    const selfDetails = {
      employeeId: props.userEmail || '',
      employeeName: props.userName || '',
      email: props.userEmail || '',
      department: 'General'
    };
    sessionStorage.setItem('loggedInEmployee', JSON.stringify(selfDetails));
    setEmployee(selfDetails);
    setSearchEmpName(props.userName || '');
  };

  // Override standard SharePoint context props with the custom employee details
  const customProps = {
    ...props,
    userName: employee.employeeName,
    userEmail: employee.email,
    employeeId: employee.employeeId,
    department: employee.department
  };

  return (
    <div className={styles.employeePanel}>
      <div className={styles.mainContent}>
        
        {/* Portal Header with User Persona and Manual Search */}
        <Stack horizontal wrap horizontalAlign="space-between" verticalAlign="center" tokens={{ childrenGap: 15 }} style={{ borderBottom: '1px solid #e2e8f0', paddingBottom: '15px' }}>
          <Persona
            text={employee.employeeName}
            secondaryText={`ID: ${employee.employeeId} | ${employee.department}`}
            size={PersonaSize.size40}
          />
          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} wrap>
            <TextField
              placeholder="Enter Employee Name..."
              styles={{ root: { width: '220px' } }}
              iconProps={{ iconName: 'Search' }}
              value={searchEmpName}
              onChange={(e, val) => setSearchEmpName(val || '')}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  handleEmployeeLookup(searchEmpName);
                }
              }}
            />
            <PrimaryButton
              text="Load Assets"
              onClick={() => handleEmployeeLookup(searchEmpName)}
              styles={{
                root: {
                  borderRadius: '8px',
                  height: '32px'
                }
              }}
            />
            {((employee.email || '').toLowerCase() !== (props.userEmail || '').toLowerCase() || 
              (employee.employeeName || '').toLowerCase() !== (props.userName || '').toLowerCase()) && (
              <DefaultButton
                text="Reset to Me"
                iconProps={{ iconName: 'Refresh' }}
                onClick={handleResetToSelf}
                styles={{
                  root: {
                    borderRadius: '8px',
                    height: '32px'
                  }
                }}
              />
            )}
          </Stack>
        </Stack>

        <div className={styles.heroSection}>
          <div className={styles.heroText}>
            <h2>Employee Management</h2>
            <p>Welcome back, {customProps.userName}!</p>
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
                <Dashboard {...customProps} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>
            
            <PivotItem headerText="Request Asset" itemIcon="ShoppingCart" itemKey="request-asset">
              <div style={{ marginTop: '20px' }}>
                <AssetRequestModule {...customProps} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="Raise Incident" itemIcon="AlertSolid" itemKey="raise-incident">
              <div style={{ marginTop: '20px' }}>
                <IncidentRequestModule {...customProps} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="My Requests" itemIcon="ReviewSolid" itemKey="my-requests">
              <div style={{ marginTop: '20px' }}>
                <MyRequests {...customProps} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="My Assigned Assets" itemIcon="CheckMark" itemKey="my-assets">
              <div style={{ marginTop: '20px' }}>
                <MyAssignedAssets {...customProps} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="Incident History" itemIcon="History" itemKey="incident-history">
              <div style={{ marginTop: '20px' }}>
                <IncidentHistory {...customProps} setIsLoading={setIsLoading} />
              </div>
            </PivotItem>

            <PivotItem headerText="Admin Approvals" itemIcon="EntitlementPolicy" itemKey="admin-approvals">
              <div style={{ marginTop: '20px' }}>
                <AdminApprovals {...customProps} setIsLoading={setIsLoading} />
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

