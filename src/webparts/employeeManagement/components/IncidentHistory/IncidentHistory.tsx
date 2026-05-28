import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  DetailsList,
  DetailsListLayoutMode,
  SelectionMode,
  IColumn,
  Icon,
  SearchBox,
  Dropdown,
  IDropdownOption,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
} from '@fluentui/react';
import { Card, ICardTokens } from '@uifabric/react-cards';
import styles from './IncidentHistory.module.scss';
import { IEmployeeManagementProps } from '../IEmployeeManagementProps';
import { InventoryService } from '../../services/InventoryService';

interface IIncidentHistoryItem {
  id: string;
  incidentId: string;
  assetId: string;
  assetName: string;
  issueType: string;
  issueDescription: string;
  priority: string;
  status: string;
  reportedDate: string;
  resolvedDate?: string;
  assignedTo?: string;
  resolution?: string;
}

const cardTokens: ICardTokens = { childrenMargin: 12 };

const statusBadgeStyles: { [key: string]: { backgroundColor: string; color: string } } = {
  Open: { backgroundColor: '#e74c3c', color: '#fff' },
  'In Progress': { backgroundColor: '#f39c12', color: '#fff' },
  Resolved: { backgroundColor: '#27ae60', color: '#fff' },
  Closed: { backgroundColor: '#666', color: '#fff' },
};

export const IncidentHistory: React.FC<IEmployeeManagementProps & { setIsLoading: (loading: boolean) => void }> = (props) => {
  const [incidents, setIncidents] = useState<IIncidentHistoryItem[]>([]);
  const [filteredIncidents, setFilteredIncidents] = useState<IIncidentHistoryItem[]>([]);
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<string | null>(null);
  const [selectedIncident, setSelectedIncident] = useState<IIncidentHistoryItem | null>(null);
  const [showDetailDialog, setShowDetailDialog] = useState(false);

  useEffect(() => {
    loadIncidents();
  }, []);

  useEffect(() => {
    filterIncidents();
  }, [searchText, statusFilter, incidents]);

  const loadIncidents = async () => {
    try {
      props.setIsLoading(true);
      const service = new InventoryService(props.apiBaseUrl);
      const data = await service.getEmployeeIncidentHistory(props.userEmail);
      setIncidents(data);
    } catch (error) {
      console.error('Error loading incident history:', error);
    } finally {
      props.setIsLoading(false);
    }
  };

  const filterIncidents = () => {
    let filtered = [...incidents];

    if (searchText) {
      filtered = filtered.filter(
        (incident) =>
          incident.assetName.toLowerCase().includes(searchText.toLowerCase()) ||
          incident.issueType.toLowerCase().includes(searchText.toLowerCase()) ||
          incident.incidentId.toLowerCase().includes(searchText.toLowerCase())
      );
    }

    if (statusFilter) {
      filtered = filtered.filter((incident) => incident.status === statusFilter);
    }

    setFilteredIncidents(filtered);
  };

  const handleViewDetails = (item: IIncidentHistoryItem) => {
    setSelectedIncident(item);
    setShowDetailDialog(true);
  };

  const handleDownloadReport = (incident: IIncidentHistoryItem) => {
    const report = `
Incident Report
==========================
Incident ID: ${incident.incidentId}
Asset: ${incident.assetName}
Issue Type: ${incident.issueType}
Priority: ${incident.priority}
Status: ${incident.status}
Reported Date: ${new Date(incident.reportedDate).toLocaleString()}
${incident.resolvedDate ? `Resolved Date: ${new Date(incident.resolvedDate).toLocaleString()}` : ''}
${incident.assignedTo ? `Assigned To: ${incident.assignedTo}` : ''}

Issue Description:
${incident.issueDescription}

${incident.resolution ? `Resolution:\n${incident.resolution}` : ''}

Generated: ${new Date().toLocaleString()}
    `;
    const element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(report));
    element.setAttribute('download', `incident-${incident.incidentId}.txt`);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  const columns: IColumn[] = [
    {
      key: 'incidentId',
      name: 'Incident ID',
      minWidth: 100,
      onRender: (item: IIncidentHistoryItem) => <Text>{item.incidentId}</Text>,
    },
    {
      key: 'assetName',
      name: 'Asset',
      minWidth: 120,
      onRender: (item: IIncidentHistoryItem) => <Text>{item.assetName}</Text>,
    },
    {
      key: 'issueType',
      name: 'Issue Type',
      minWidth: 120,
      onRender: (item: IIncidentHistoryItem) => <Text>{item.issueType}</Text>,
    },
    {
      key: 'priority',
      name: 'Priority',
      minWidth: 100,
      onRender: (item: IIncidentHistoryItem) => (
        <Text style={{ fontWeight: 600, color: item.priority === 'Critical' ? '#e74c3c' : item.priority === 'High' ? '#f39c12' : '#0078d4' }}>
          {item.priority}
        </Text>
      ),
    },
    {
      key: 'status',
      name: 'Status',
      minWidth: 120,
      onRender: (item: IIncidentHistoryItem) => (
        <div
          style={{
            ...statusBadgeStyles[item.status],
            padding: '6px 12px',
            borderRadius: '4px',
            textAlign: 'center',
            fontWeight: 600,
            fontSize: '12px',
          }}
        >
          {item.status}
        </div>
      ),
    },
    {
      key: 'reportedDate',
      name: 'Reported',
      minWidth: 120,
      onRender: (item: IIncidentHistoryItem) => (
        <Text>{new Date(item.reportedDate).toLocaleDateString()}</Text>
      ),
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 150,
      onRender: (item: IIncidentHistoryItem) => (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text="View"
            onClick={() => handleViewDetails(item)}
            styles={{
              root: { padding: '4px 12px', fontSize: '12px', height: '24px' },
            }}
          />
          <PrimaryButton
            text="Download"
            onClick={() => handleDownloadReport(item)}
            styles={{
              root: { padding: '4px 12px', fontSize: '12px', height: '24px' },
            }}
          />
        </Stack>
      ),
    },
  ];

  const statusFilterOptions: IDropdownOption[] = [
    { key: '', text: 'All Status' },
    { key: 'Open', text: 'Open' },
    { key: 'In Progress', text: 'In Progress' },
    { key: 'Resolved', text: 'Resolved' },
    { key: 'Closed', text: 'Closed' },
  ];

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Text variant="xLarge" block style={{ fontWeight: 600 }}>
        Incident History
      </Text>

      <Card>
        <Card.Section tokens={cardTokens}>
          <Stack tokens={{ childrenGap: 15 }}>
            {/* Filters */}
            <Stack horizontal tokens={{ childrenGap: 15 }} wrap>
              <SearchBox
                placeholder="Search by incident ID, asset name, or issue type..."
                value={searchText}
                onChange={(ev, newValue) => setSearchText(newValue || '')}
                style={{ flex: 1, minWidth: '250px' }}
              />
              <Dropdown
                placeholder="Filter by status"
                options={statusFilterOptions}
                onChange={(ev, option) => setStatusFilter(option?.key as string | null)}
                style={{ width: '200px' }}
              />
            </Stack>

            {/* Items Count */}
            <Text variant="small" style={{ color: '#666' }}>
              Showing {filteredIncidents.length} of {incidents.length} incidents
            </Text>

            {/* Details List */}
            {filteredIncidents.length > 0 ? (
              <DetailsList
                items={filteredIncidents}
                columns={columns}
                setKey="set-items"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            ) : (
              <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '300px' }}>
                <Icon iconName="ClearFilter" style={{ fontSize: '48px', color: '#ccc', marginBottom: '10px' }} />
                <Text variant="large" style={{ color: '#666' }}>
                  No incidents found.
                </Text>
              </Stack>
            )}
          </Stack>
        </Card.Section>
      </Card>

      {/* Detail Dialog */}
      <Dialog
        hidden={!showDetailDialog}
        onDismiss={() => setShowDetailDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Incident Details',
          closeButtonAriaLabel: 'Close',
        }}
        minWidth={600}
      >
        {selectedIncident && (
          <Stack tokens={{ childrenGap: 15 }}>
            <TextField label="Incident ID" value={selectedIncident.incidentId} disabled />
            <TextField label="Asset Name" value={selectedIncident.assetName} disabled />
            <TextField label="Issue Type" value={selectedIncident.issueType} disabled />
            <TextField label="Priority" value={selectedIncident.priority} disabled />
            <TextField label="Status" value={selectedIncident.status} disabled />
            <TextField
              label="Issue Description"
              value={selectedIncident.issueDescription}
              multiline
              rows={4}
              disabled
            />
            {selectedIncident.resolution && (
              <TextField
                label="Resolution"
                value={selectedIncident.resolution}
                multiline
                rows={4}
                disabled
              />
            )}
            <TextField
              label="Reported Date"
              value={new Date(selectedIncident.reportedDate).toLocaleString()}
              disabled
            />
            {selectedIncident.resolvedDate && (
              <TextField
                label="Resolved Date"
                value={new Date(selectedIncident.resolvedDate).toLocaleString()}
                disabled
              />
            )}
            {selectedIncident.assignedTo && (
              <TextField label="Assigned To" value={selectedIncident.assignedTo} disabled />
            )}
          </Stack>
        )}
        <DialogFooter>
          <PrimaryButton text="Close" onClick={() => setShowDetailDialog(false)} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

