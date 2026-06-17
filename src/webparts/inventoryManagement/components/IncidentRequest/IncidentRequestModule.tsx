import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  TextField,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  DefaultButton,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { Card, ICardTokens } from '@uifabric/react-cards';
import styles from './IncidentRequestModule.module.scss';
import { IInventoryManagementProps } from '../../models/IInventoryManagementProps';
import { IncidentService } from '../../services/IncidentService';

interface IIncidentRequestModuleProps extends IInventoryManagementProps {
  employeeId?: string;
  department?: string;
  setIsLoading: (loading: boolean) => void;
}

interface IIncidentForm {
  employeeName: string;
  employeeId: string;
  employeeEmail?: string;
  serialNo: string;
  assetName: string;
  incidentType: string;
  priority: string;
  description: string;
  raisedDate: string;
  status: string;
  raisedTo: string;
  assignedDate: string;
}

const incidentTypeOptions: IDropdownOption[] = [
  { key: 'Hardware Issue', text: 'Hardware Issue' },
  { key: 'Software Issue', text: 'Software Issue' },
  { key: 'Network Issue', text: 'Network Issue' },
  { key: 'Asset Damage', text: 'Asset Damage' },
  { key: 'Replacement Request', text: 'Replacement Request' },
  { key: 'Access Issue', text: 'Access Issue' },
  { key: 'Login Issue', text: 'Login Issue' },
  { key: 'Performance Issue', text: 'Performance Issue' },
  { key: 'Email Issue', text: 'Email Issue' },
  { key: 'Printer Issue', text: 'Printer Issue' },
  { key: 'Other', text: 'Other' }
];

const priorityOptions: IDropdownOption[] = [
  { key: 'Low', text: 'Low' },
  { key: 'Medium', text: 'Medium' },
  { key: 'High', text: 'High' },
  { key: 'Critical', text: 'Critical' },
];

const raisedToOptions: IDropdownOption[] = [
  { key: 'Admin', text: 'Admin' }
];

const cardTokens: ICardTokens = { childrenMargin: 12 };

export const IncidentRequestModule: React.FC<IIncidentRequestModuleProps> = (props) => {
  const [formData, setFormData] = useState<IIncidentForm>({
    employeeName: props.userDisplayName || '',
    employeeId: props.employeeId || '',
    employeeEmail: props.userEmail || '',
    serialNo: '',
    assetName: '',
    incidentType: '',
    priority: 'Medium',
    description: '',
    raisedDate: new Date().toLocaleString(),
    status: 'Open',
    raisedTo: 'Admin',
    assignedDate: '',
  });

  const [assignedAssets, setAssignedAssets] = useState<any[]>([]);
  const [isLoadingAssets, setIsLoadingAssets] = useState(true);
  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);
  const [isSubmitting, setIsSubmitting] = useState(false);

  useEffect(() => {
    const timer = setTimeout(() => {
      loadAssignedAssetsForName(formData.employeeName);
    }, 500);
    return () => clearTimeout(timer);
  }, [formData.employeeName]);

  // Sync with props when employee context changes
  useEffect(() => {
    setFormData((prev) => ({
      ...prev,
      employeeName: props.userDisplayName || '',
      employeeId: props.employeeId || '',
      employeeEmail: props.userEmail || '',
    }));
  }, [props.userDisplayName, props.employeeId, props.userEmail]);

  const loadAssignedAssetsForName = async (name: string) => {
    if (!name.trim()) {
      setAssignedAssets([]);
      setIsLoadingAssets(false);
      return;
    }
    try {
      setIsLoadingAssets(true);
      const service = new IncidentService(props.spContext);
      const details = await service.getEmployeeDetailsByName(name);
      
      setFormData(prev => ({
        ...prev,
        employeeId: details.employeeId || prev.employeeId,
        employeeEmail: details.email || prev.employeeEmail
      }));

      const assets = await service.getEmployeeAssignedAssets(details.email || details.employeeName);
      setAssignedAssets(assets);
    } catch (error) {
      console.error('Error loading assigned assets:', error);
      setAssignedAssets([]);
    } finally {
      setIsLoadingAssets(false);
    }
  };

  const handleInputChange = (fieldName: string, value: any) => {
    setFormData((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const handleSubmit = async () => {
    try {
      if (!formData.incidentType || !formData.description) {
        setMessage({ type: MessageBarType.error, text: 'Please fill in all required fields.' });
        return;
      }

      setIsSubmitting(true);
      const service = new IncidentService(props.spContext);
      
      const payload = {
        ...formData,
        employeeEmail: formData.employeeEmail || props.userEmail,
        employeeName: formData.employeeName,
        employeeId: formData.employeeId,
      };
      
      console.log('Submitting incident payload:', payload);
      await service.createIncidentRequest(payload);

      setMessage({ type: MessageBarType.success, text: 'Incident reported successfully!' });

      setTimeout(() => {
        setFormData({
          employeeName: props.userDisplayName || '',
          employeeId: props.employeeId || '',
          employeeEmail: props.userEmail || '',
          serialNo: '',
          assetName: '',
          incidentType: '',
          priority: 'Medium',
          description: '',
          raisedDate: new Date().toLocaleString(),
          status: 'Open',
          raisedTo: 'Admin',
          assignedDate: '',
        });
        setMessage(null);
        // Refresh assigned assets in case it changes
        loadAssignedAssetsForName(props.userDisplayName);
      }, 2000);
    } catch (error) {
      console.error('Error submitting incident:', error);
      const errorMessage = error instanceof Error ? error.message : 'Failed to report incident. Please try again.';
      setMessage({ type: MessageBarType.error, text: `Error: ${errorMessage}` });
    } finally {
      setIsSubmitting(false);
    }
  };

  const handleCancel = () => {
    setFormData({
      employeeName: props.userDisplayName || '',
      employeeId: props.employeeId || '',
      employeeEmail: props.userEmail || '',
      serialNo: '',
      assetName: '',
      incidentType: '',
      priority: 'Medium',
      description: '',
      raisedDate: new Date().toLocaleString(),
      status: 'Open',
      raisedTo: 'Admin',
      assignedDate: '',
    });
    setMessage(null);
  };

  const assetOptions: IDropdownOption[] = assignedAssets.map(a => ({
    key: a.id,
    text: `${a.assetName} (S/N: ${a.serialNumber || 'N/A'})`,
  }));

  const selectedAssetKey = assignedAssets.find(a => a.serialNumber === formData.serialNo && a.assetName === formData.assetName)?.id;

  return (
    <div className={styles.incidentRequestModule}>
      <Stack tokens={{ childrenGap: 20 }}>
        <Text variant="xLarge" block style={{ fontWeight: 600 }}>
          Raise Incident
        </Text>

        {message && (
          <MessageBar messageBarType={message.type} isMultiline>
            {message.text}
          </MessageBar>
        )}

        {!isLoadingAssets && assignedAssets.length === 0 && (
          <MessageBar messageBarType={MessageBarType.info}>
            You currently have no assets assigned. You can still raise generic incidents.
          </MessageBar>
        )}

        <Card>
          <Card.Section tokens={cardTokens}>
            <Stack tokens={{ childrenGap: 15 }}>
              <Stack tokens={{ childrenGap: 10 }}>
                <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                  Employee Information
                </Text>
                <TextField 
                  label="Employee Name *" 
                  value={formData.employeeName} 
                  onChange={(ev, val) => handleInputChange('employeeName', val || '')}
                  required
                />
              </Stack>

              <Stack tokens={{ childrenGap: 10 }}>
                <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                  Asset Details
                </Text>

                <Dropdown
                  label="Select Assigned Asset"
                  placeholder={isLoadingAssets ? "Loading assets..." : "Choose one of your assigned assets"}
                  options={assetOptions}
                  selectedKey={selectedAssetKey}
                  onChange={(ev, option) => {
                    const selected = assignedAssets.find(a => a.id === option?.key);
                    if (selected) {
                      handleInputChange('assetName', selected.assetName);
                      handleInputChange('serialNo', selected.serialNumber);
                      handleInputChange('assignedDate', selected.assignmentDate);
                    } else {
                      handleInputChange('assetName', '');
                      handleInputChange('serialNo', '');
                      handleInputChange('assignedDate', '');
                    }
                  }}
                  disabled={isLoadingAssets}
                />

                {formData.assetName && (
                  <TextField
                    label="Asset Name"
                    value={formData.assetName}
                    disabled
                  />
                )}

                {formData.serialNo && (
                  <TextField
                    label="Serial NO"
                    value={formData.serialNo}
                    disabled
                  />
                )}

                <Dropdown
                  label="Incident Type *"
                  options={incidentTypeOptions}
                  selectedKey={formData.incidentType}
                  onChange={(ev, option) => handleInputChange('incidentType', option?.key)}
                  required
                  placeholder="Select Incident Type"
                />

                <Dropdown
                  label="Priority"
                  options={priorityOptions}
                  selectedKey={formData.priority}
                  onChange={(ev, option) => handleInputChange('priority', option?.key)}
                />

                <TextField
                  label="Description *"
                  multiline
                  rows={5}
                  placeholder="Describe the issue..."
                  value={formData.description}
                  onChange={(ev, newValue) => handleInputChange('description', newValue)}
                  required
                />

                <Dropdown
                  label="Raised To"
                  options={raisedToOptions}
                  selectedKey={formData.raisedTo}
                  onChange={(ev, option) => handleInputChange('raisedTo', option?.key)}
                  placeholder="Select Team"
                />
              </Stack>

              <Stack tokens={{ childrenGap: 10 }}>
                <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                  Timeline & Status
                </Text>
                <TextField
                  label="Raised Date"
                  value={formData.raisedDate}
                  readOnly
                />
                <TextField
                  label="Status"
                  value={formData.status}
                  readOnly
                />
              </Stack>

              <Stack horizontal tokens={{ childrenGap: 10 }}>
                <PrimaryButton
                  text="Report Incident"
                  onClick={handleSubmit}
                  disabled={isSubmitting}
                />
                <DefaultButton
                  text="Cancel"
                  onClick={handleCancel}
                />
              </Stack>
            </Stack>
          </Card.Section>
        </Card>
      </Stack>
    </div>
  );
};
