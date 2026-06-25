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
  TextField,
  Panel,
  PanelType,
} from '@fluentui/react';
import { jsPDF } from 'jspdf';
import styles from './IncidentHistory.module.scss';
import { IInventoryManagementProps } from '../../models/IInventoryManagementProps';
import { IncidentService } from '../../services/IncidentService';

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

export const IncidentHistory: React.FC<IInventoryManagementProps & { setIsLoading: (loading: boolean) => void; userRole?: string; }> = (props) => {
  const [incidents, setIncidents] = useState<IIncidentHistoryItem[]>([]);
  const [filteredIncidents, setFilteredIncidents] = useState<IIncidentHistoryItem[]>([]);
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<string | null>(null);
  const [selectedIncident, setSelectedIncident] = useState<IIncidentHistoryItem | null>(null);
  const [showDetailPanel, setShowDetailPanel] = useState(false);
  const [tempResolution, setTempResolution] = useState('');

  const getPriorityBadgeStyle = (priority?: string) => {
    const p = priority || 'Medium';
    let backgroundColor = '#f3f4f6';
    let color = '#4b5563';
    if (p === 'High' || p === 'Critical') {
      backgroundColor = '#fee2e2';
      color = '#b91c1c';
    } else if (p === 'Low') {
      backgroundColor = '#dbeafe';
      color = '#1e3a8a';
    }
    return {
      backgroundColor,
      color,
      padding: '4px 10px',
      borderRadius: '9999px',
      fontSize: '0.75rem',
      fontWeight: 600 as const,
      display: 'inline-block'
    };
  };

  const getStatusBadgeStyle = (status?: string) => {
    const s = status || 'Open';
    let backgroundColor = '#fee2e2';
    let color = '#991b1b';
    if (s === 'In Progress') {
      backgroundColor = '#fef3c7';
      color = '#92400e';
    } else if (s === 'Resolved') {
      backgroundColor = '#dcfce7';
      color = '#166534';
    } else if (s === 'Closed') {
      backgroundColor = '#f3f4f6';
      color = '#4b5563';
    }
    return {
      backgroundColor,
      color,
      padding: '4px 12px',
      borderRadius: '9999px',
      fontSize: '0.75rem',
      fontWeight: 600 as const,
      display: 'inline-block',
      textAlign: 'center' as const
    };
  };

  useEffect(() => {
    loadIncidents();
  }, [props.userEmail]);

  useEffect(() => {
    filterIncidents();
  }, [searchText, statusFilter, incidents]);

  const loadIncidents = async () => {
    try {
      props.setIsLoading(true);
      const service = new IncidentService(props.spContext);
      const isAdmin = props.userRole === 'Admin';
      const data = await service.getEmployeeIncidentHistory(props.userEmail, isAdmin);
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
          (incident.assetName || '').toLowerCase().includes(searchText.toLowerCase()) ||
          (incident.issueType || '').toLowerCase().includes(searchText.toLowerCase()) ||
          (incident.incidentId || '').toLowerCase().includes(searchText.toLowerCase())
      );
    }

    if (statusFilter) {
      filtered = filtered.filter((incident) => incident.status === statusFilter);
    }

    setFilteredIncidents(filtered);
  };

  const handleViewDetails = (item: IIncidentHistoryItem) => {
    setSelectedIncident(item);
    setTempResolution(item.resolution || '');
    setShowDetailPanel(true);
  };

  const handleStatusChange = async (incident: IIncidentHistoryItem, newStatus: string) => {
    try {
      props.setIsLoading(true);
      const service = new IncidentService(props.spContext);
      await service.updateIncidentStatus(incident.id, newStatus, incident.resolution);
      
      const updatedIncident = {
        ...incident,
        status: newStatus,
        resolvedDate: newStatus === 'Resolved' || newStatus === 'Closed' ? new Date().toISOString() : incident.resolvedDate
      };
      setSelectedIncident(updatedIncident);
      await loadIncidents();
    } catch (error) {
      console.error('Error updating status:', error);
    } finally {
      props.setIsLoading(false);
    }
  };

  const handleSaveResolution = async (incident: IIncidentHistoryItem) => {
    try {
      props.setIsLoading(true);
      const service = new IncidentService(props.spContext);
      await service.updateIncidentStatus(incident.id, incident.status, tempResolution);
      
      const updatedIncident = {
        ...incident,
        resolution: tempResolution
      };
      setSelectedIncident(updatedIncident);
      await loadIncidents();
    } catch (error) {
      console.error('Error saving resolution:', error);
    } finally {
      props.setIsLoading(false);
    }
  };

  const handleDownloadReport = (incident: IIncidentHistoryItem) => {
    try {
      const doc = new jsPDF();
      
      // Top header banner
      doc.setFillColor(0, 90, 158); // #005a9e (Deep blue theme color)
      doc.rect(0, 0, 210, 25, 'F');
      
      doc.setTextColor(255, 255, 255);
      doc.setFont("helvetica", "bold");
      doc.setFontSize(16);
      doc.text("MSFT INVENTORY MANAGEMENT", 14, 16);
      
      // Document Title
      doc.setTextColor(51, 65, 85); // Slate 700
      doc.setFontSize(14);
      doc.text("INCIDENT REPORT", 14, 38);
      
      // Metadata
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(`Generated on: ${new Date().toLocaleString()}`, 14, 44);
      
      // Separator line
      doc.setDrawColor(226, 232, 240); // Slate 200
      doc.line(14, 48, 196, 48);
      
      // Specifications Section Title
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11);
      doc.text("INCIDENT SPECIFICATIONS", 14, 58);
      
      // Render Specifications Key-Value grid
      let y = 68;
      const printField = (label: string, value: string) => {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(9);
        doc.setTextColor(100, 116, 139); // Slate 500
        doc.text(label, 14, y);
        
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9.5);
        doc.setTextColor(15, 23, 42); // Slate 900
        doc.text(value, 55, y);
        y += 8;
      };
      
      printField("Incident ID:", incident.incidentId);
      printField("Asset Name:", incident.assetName);
      printField("Issue Type:", incident.issueType);
      printField("Priority:", incident.priority || "Medium");
      printField("Current Status:", incident.status || "Open");
      printField("Reported Date:", new Date(incident.reportedDate).toLocaleString());
      
      if (incident.assignedTo) {
        printField("Assigned To:", incident.assignedTo);
      }
      if (incident.resolvedDate) {
        printField("Resolved Date:", new Date(incident.resolvedDate).toLocaleString());
      }
      
      // Issue Description Title
      y += 4;
      doc.setFont("helvetica", "bold");
      doc.setFontSize(11);
      doc.setTextColor(51, 65, 85);
      doc.text("ISSUE DESCRIPTION", 14, y);
      y += 6;
      
      // Issue Description Box
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9.5);
      doc.setTextColor(51, 65, 85);
      
      const splitDesc = doc.splitTextToSize(incident.issueDescription || "No description provided.", 170);
      const descHeight = splitDesc.length * 6 + 10;
      
      // Draw background box
      doc.setFillColor(248, 250, 252); // slate 50
      doc.setDrawColor(226, 232, 240); // slate 200
      doc.rect(14, y, 182, descHeight, 'FD');
      
      // Draw left accent bar
      doc.setFillColor(100, 116, 139); // slate 500
      doc.rect(14, y, 3, descHeight, 'F');
      
      // Draw text
      let textY = y + 8;
      splitDesc.forEach((line: string) => {
        doc.text(line, 22, textY);
        textY += 6;
      });
      
      y += descHeight + 10;
      
      // Resolution Details (if resolved/closed)
      if (incident.resolution) {
        doc.setFont("helvetica", "bold");
        doc.setFontSize(11);
        doc.setTextColor(51, 65, 85);
        doc.text("RESOLUTION SUMMARY", 14, y);
        y += 6;
        
        doc.setFont("helvetica", "normal");
        doc.setFontSize(9.5);
        doc.setTextColor(22, 101, 52); // green 800
        
        const splitRes = doc.splitTextToSize(incident.resolution, 170);
        const resHeight = splitRes.length * 6 + 10;
        
        // Draw green background box
        doc.setFillColor(240, 253, 244); // green 50
        doc.setDrawColor(220, 252, 231); // green 200
        doc.rect(14, y, 182, resHeight, 'FD');
        
        // Draw green left accent bar
        doc.setFillColor(22, 101, 52); // green 800
        doc.rect(14, y, 3, resHeight, 'F');
        
        // Draw resolution text
        let resTextY = y + 8;
        splitRes.forEach((line: string) => {
          doc.text(line, 22, resTextY);
          resTextY += 6;
        });
      }
      
      doc.save(`incident-${incident.incidentId}.pdf`);
    } catch (error) {
      console.error('Error generating PDF report:', error);
    }
  };

  const columns: IColumn[] = [
    {
      key: 'incidentId',
      name: 'Incident ID',
      fieldName: 'incidentId',
      minWidth: 90,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IIncidentHistoryItem) => <Text>{item.incidentId}</Text>,
    },
    {
      key: 'assetName',
      name: 'Asset',
      fieldName: 'assetName',
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IIncidentHistoryItem) => <Text>{item.assetName}</Text>,
    },
    {
      key: 'issueType',
      name: 'Issue Type',
      fieldName: 'issueType',
      minWidth: 100,
      maxWidth: 130,
      isResizable: true,
      onRender: (item: IIncidentHistoryItem) => <Text>{item.issueType}</Text>,
    },
    {
      key: 'priority',
      name: 'Priority',
      fieldName: 'priority',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IIncidentHistoryItem) => {
        return (
          <span style={getPriorityBadgeStyle(item.priority)}>
            {item.priority || 'Medium'}
          </span>
        );
      },
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'status',
      minWidth: 90,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IIncidentHistoryItem) => {
        return (
          <span style={getStatusBadgeStyle(item.status)}>
            {item.status || 'Open'}
          </span>
        );
      },
    },
    {
      key: 'reportedDate',
      name: 'Reported',
      fieldName: 'reportedDate',
      minWidth: 90,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IIncidentHistoryItem) => {
        if (!item.reportedDate) return <Text>-</Text>;
        try {
          return <Text>{new Date(item.reportedDate).toLocaleDateString()}</Text>;
        } catch {
          return <Text>{item.reportedDate}</Text>;
        }
      },
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 160,
      maxWidth: 220,
      isResizable: true,
      onRender: (item: IIncidentHistoryItem) => (
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton
            text="View"
            onClick={() => handleViewDetails(item)}
            styles={{
              root: { padding: '2px 10px', fontSize: '11px', height: '24px' },
            }}
          />
          <PrimaryButton
            text="Download"
            onClick={() => handleDownloadReport(item)}
            styles={{
              root: { padding: '2px 10px', fontSize: '11px', height: '24px' },
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
    <div style={{ marginTop: '20px' }}>
      <Stack tokens={{ childrenGap: 15 }}>
        {/* Filters */}
        <div style={{ display: 'flex', gap: '15px', alignItems: 'center', flexWrap: 'wrap', marginBottom: '5px' }}>
          <SearchBox
            placeholder="Search by incident ID, asset name, or issue type..."
            value={searchText}
            onChange={(ev, newValue) => setSearchText(newValue || '')}
            onClear={() => setSearchText('')}
            styles={{ root: { width: '100%', maxWidth: 400 } }}
          />
          <Dropdown
            placeholder="Filter by status"
            options={statusFilterOptions}
            onChange={(ev, option) => setStatusFilter(option?.key as string | null || null)}
            styles={{ root: { width: 200 } }}
          />
        </div>

        {/* Items Count */}
        <Text variant="small" style={{ color: 'var(--text-muted, #6b7280)', display: 'block' }}>
          Showing {filteredIncidents.length} of {incidents.length} incidents
        </Text>

        {/* Details List */}
        {filteredIncidents.length > 0 ? (
          <DetailsList
            items={filteredIncidents}
            columns={columns}
            setKey="incident-list"
            layoutMode={DetailsListLayoutMode.justified}
            selectionMode={SelectionMode.none}
          />
        ) : (
          <div style={{
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center',
            minHeight: '250px',
            border: '1px dashed #e5e7eb',
            borderRadius: '8px',
            padding: '30px'
          }}>
            <Icon iconName="ClearFilter" style={{ fontSize: '36px', color: '#9ca3af', marginBottom: '10px' }} />
            <Text variant="medium" style={{ color: '#6b7280' }}>
              No incidents found.
            </Text>
          </div>
        )}
      </Stack>

      {/* Detail Panel */}
      <Panel
        isOpen={showDetailPanel}
        onDismiss={() => setShowDetailPanel(false)}
        type={PanelType.medium}
        headerText="Incident Details"
        closeButtonAriaLabel="Close"
      >
        {selectedIncident && (
          <div style={{ marginTop: '10px' }}>
            <p style={{ color: '#6b7280', fontSize: '0.88rem', margin: '0 0 20px 0' }}>
              <strong>Reported:</strong> {new Date(selectedIncident.reportedDate).toLocaleString()}
            </p>

            <div style={{ padding: '12px 15px', backgroundColor: '#f1f5f9', borderRadius: '6px', marginBottom: '20px', borderLeft: '4px solid #64748b' }}>
              <p style={{ margin: 0, fontSize: '0.92rem', color: '#334155', lineHeight: '1.5', whiteSpace: 'pre-wrap' }}>
                {selectedIncident.issueDescription}
              </p>
            </div>

            <div style={{ backgroundColor: '#ffffff', padding: '16px', borderRadius: '8px', border: '1px solid #e5e7eb', marginBottom: '20px' }}>
              <h4 style={{ margin: '0 0 12px 0', color: '#111827', fontSize: '1rem', borderBottom: '1px solid #f3f4f6', paddingBottom: '8px' }}>
                Incident Specifications
              </h4>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '0.88rem', alignItems: 'center' }}>
                <div><span style={{ color: '#6b7280' }}>Incident ID:</span> <strong style={{ color: '#111827' }}>{selectedIncident.incidentId}</strong></div>
                <div><span style={{ color: '#6b7280' }}>Asset Name:</span> <strong style={{ color: '#111827' }}>{selectedIncident.assetName}</strong></div>
                <div><span style={{ color: '#6b7280' }}>Issue Type:</span> <strong style={{ color: '#111827' }}>{selectedIncident.issueType}</strong></div>
                
                <div>
                  <span style={{ color: '#6b7280', marginRight: '6px' }}>Priority:</span>
                  <span style={getPriorityBadgeStyle(selectedIncident.priority)}>
                    {selectedIncident.priority || 'Medium'}
                  </span>
                </div>
                
                {props.userRole === 'Admin' ? (
                  <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span style={{ color: '#6b7280' }}>Status:</span>
                    <Dropdown
                      selectedKey={selectedIncident.status || 'Open'}
                      options={[
                        { key: 'Open', text: 'Open' },
                        { key: 'In Progress', text: 'In Progress' },
                        { key: 'Resolved', text: 'Resolved' },
                        { key: 'Closed', text: 'Closed' }
                      ]}
                      onChange={(ev, option) => handleStatusChange(selectedIncident, option?.key as string)}
                      styles={{ root: { width: 120 } }}
                    />
                  </div>
                ) : (
                  <div>
                    <span style={{ color: '#6b7280', marginRight: '6px' }}>Status:</span>
                    <span style={getStatusBadgeStyle(selectedIncident.status)}>
                      {selectedIncident.status || 'Open'}
                    </span>
                  </div>
                )}

                {selectedIncident.assignedTo && (
                  <div><span style={{ color: '#6b7280' }}>Assigned To:</span> <strong style={{ color: '#111827' }}>{selectedIncident.assignedTo}</strong></div>
                )}
              </div>
            </div>

            {props.userRole === 'Admin' && (selectedIncident.status === 'Resolved' || selectedIncident.status === 'Closed') ? (
              <div style={{ backgroundColor: '#f8fafc', padding: '16px', borderRadius: '8px', border: '1px solid #e2e8f0', marginBottom: '20px' }}>
                <h4 style={{ margin: '0 0 12px 0', color: '#1e293b', fontSize: '1rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                  <Icon iconName="CheckMark" style={{ color: '#166534', fontWeight: 'bold' }} /> Update Resolution Details
                </h4>
                <Stack tokens={{ childrenGap: 10 }}>
                  {selectedIncident.resolvedDate && (
                    <div style={{ fontSize: '0.88rem' }}>
                      <span style={{ color: '#6b7280' }}>Resolved Date:</span>{' '}
                      <strong style={{ color: '#111827' }}>{new Date(selectedIncident.resolvedDate).toLocaleString()}</strong>
                    </div>
                  )}
                  <TextField
                    label="Resolution Summary"
                    multiline
                    rows={3}
                    value={tempResolution}
                    onChange={(ev, newValue) => setTempResolution(newValue || '')}
                    placeholder="Describe how this issue was resolved..."
                  />
                  <PrimaryButton
                    text="Save Resolution"
                    onClick={() => handleSaveResolution(selectedIncident)}
                    styles={{ root: { alignSelf: 'flex-start' } }}
                  />
                </Stack>
              </div>
            ) : (
              selectedIncident.resolution && (
                <div style={{ backgroundColor: '#f8fafc', padding: '16px', borderRadius: '8px', border: '1px solid #e2e8f0', marginBottom: '20px' }}>
                  <h4 style={{ margin: '0 0 12px 0', color: '#1e293b', fontSize: '1rem', display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <Icon iconName="CheckMark" style={{ color: '#166534', fontWeight: 'bold' }} /> Resolution Details
                  </h4>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '8px', fontSize: '0.88rem' }}>
                    {selectedIncident.resolvedDate && (
                      <div>
                        <span style={{ color: '#6b7280' }}>Resolved Date:</span>{' '}
                        <strong style={{ color: '#111827' }}>{new Date(selectedIncident.resolvedDate).toLocaleString()}</strong>
                      </div>
                    )}
                    <div style={{ padding: '10px', backgroundColor: '#f0fdf4', borderRadius: '6px', border: '1px solid #dcfce7', color: '#166534', fontSize: '0.88rem', lineHeight: '1.4', whiteSpace: 'pre-wrap' }}>
                      <strong>Resolution Summary:</strong> {selectedIncident.resolution}
                    </div>
                  </div>
                </div>
              )
            )}
          </div>
        )}
      </Panel>
    </div>
  );
};
