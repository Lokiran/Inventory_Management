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
  PrimaryButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { Card, ICardTokens } from '@uifabric/react-cards';
import { IEmployeeManagementProps } from '../IEmployeeManagementProps';
import { InventoryService } from '../../services/InventoryService';

interface IPendingRequest {
  id: string;
  employeeName: string;
  employeeId: string;
  serialNo: string;
  assetName: string;
  priority: string;
  reason: string;
  requestDate: string;
  status: string;
}

const cardTokens: ICardTokens = { childrenMargin: 12 };

export const AdminApprovals: React.FC<IEmployeeManagementProps & { setIsLoading: (loading: boolean) => void }> = (props) => {
  const [requests, setRequests] = useState<IPendingRequest[]>([]);
  const [filteredRequests, setFilteredRequests] = useState<IPendingRequest[]>([]);
  const [searchText, setSearchText] = useState('');
  const [message, setMessage] = useState<{ type: MessageBarType; text: string } | null>(null);
  const [processingId, setProcessingId] = useState<string | null>(null);

  useEffect(() => {
    loadPendingRequests();
  }, []);

  useEffect(() => {
    filterRequests();
  }, [searchText, requests]);

  const loadPendingRequests = async () => {
    try {
      props.setIsLoading(true);
      const service = new InventoryService(props.spContext);
      const data = await service.getPendingRequests();
      setRequests(data);
    } catch (error) {
      console.error('Error loading pending requests:', error);
      setMessage({ type: MessageBarType.error, text: 'Failed to load pending requests.' });
    } finally {
      props.setIsLoading(false);
    }
  };

  const filterRequests = () => {
    let filtered = [...requests];
    if (searchText) {
      filtered = filtered.filter(
        (req) =>
          req.assetName.toLowerCase().includes(searchText.toLowerCase()) ||
          req.employeeName.toLowerCase().includes(searchText.toLowerCase()) ||
          req.employeeId.toLowerCase().includes(searchText.toLowerCase()) ||
          req.reason.toLowerCase().includes(searchText.toLowerCase())
      );
    }
    setFilteredRequests(filtered);
  };

  const handleApprove = async (item: IPendingRequest) => {
    try {
      setProcessingId(item.id);
      setMessage(null);
      
      const service = new InventoryService(props.spContext);
      await service.approveRequest(item.id, props.userName);
      
      setMessage({
        type: MessageBarType.success,
        text: `Successfully approved request for ${item.employeeName} (${item.assetName}). Asset link saved in MappingList.`,
      });
      
      // Reload pending items
      await loadPendingRequests();
    } catch (error: any) {
      console.error('Error approving request:', error);
      setMessage({ type: MessageBarType.error, text: `Approval failed: ${error.message || error}` });
    } finally {
      setProcessingId(null);
    }
  };

  const handleReject = async (item: IPendingRequest) => {
    try {
      setProcessingId(item.id);
      setMessage(null);
      
      const service = new InventoryService(props.spContext);
      await service.rejectRequest(item.id);
      
      setMessage({
        type: MessageBarType.warning,
        text: `Rejected request for ${item.employeeName} (${item.assetName}).`,
      });
      
      // Reload pending items
      await loadPendingRequests();
    } catch (error: any) {
      console.error('Error rejecting request:', error);
      setMessage({ type: MessageBarType.error, text: `Rejection failed: ${error.message || error}` });
    } finally {
      setProcessingId(null);
    }
  };

  const columns: IColumn[] = [
    {
      key: 'employeeName',
      name: 'Employee Name',
      fieldName: 'employeeName',
      minWidth: 120,
      isResizable: true,
    },
    {
      key: 'employeeId',
      name: 'Employee ID',
      fieldName: 'employeeId',
      minWidth: 80,
      isResizable: true,
    },
    {
      key: 'assetName',
      name: 'Requested Asset',
      fieldName: 'assetName',
      minWidth: 120,
      isResizable: true,
    },
    {
      key: 'serialNo',
      name: 'Serial Number',
      fieldName: 'serialNo',
      minWidth: 100,
      isResizable: true,
    },
    {
      key: 'priority',
      name: 'Priority',
      minWidth: 80,
      isResizable: true,
      onRender: (item: IPendingRequest) => (
        <span style={{
          padding: '4px 8px',
          borderRadius: '4px',
          backgroundColor: (item.priority || '').toLowerCase() === 'high' || (item.priority || '').toLowerCase() === 'urgent' || (item.priority || '').toLowerCase() === 'critical' ? '#d13438' : '#0078d4',
          color: 'white',
          fontSize: '12px',
          fontWeight: 600,
        }}>
          {item.priority}
        </span>
      ),
    },
    {
      key: 'reason',
      name: 'Reason',
      fieldName: 'reason',
      minWidth: 150,
      isResizable: true,
    },
    {
      key: 'requestDate',
      name: 'Requested Date',
      minWidth: 100,
      isResizable: true,
      onRender: (item: IPendingRequest) => (
        <Text>{item.requestDate ? new Date(item.requestDate).toLocaleDateString() : 'N/A'}</Text>
      ),
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 180,
      onRender: (item: IPendingRequest) => {
        if (processingId === item.id) {
          return <Spinner size={SpinnerSize.small} label="Processing..." />;
        }
        return (
          <Stack horizontal tokens={{ childrenGap: 10 }}>
            <PrimaryButton
              text="Approve"
              onClick={() => handleApprove(item)}
              styles={{
                root: { padding: '4px 12px', fontSize: '12px', height: '24px', backgroundColor: '#107c10', borderColor: '#107c10' },
                rootHovered: { backgroundColor: '#0b590b', borderColor: '#0b590b' },
              }}
            />
            <PrimaryButton
              text="Reject"
              onClick={() => handleReject(item)}
              styles={{
                root: { padding: '4px 12px', fontSize: '12px', height: '24px', backgroundColor: '#a4262c', borderColor: '#a4262c' },
                rootHovered: { backgroundColor: '#751b1f', borderColor: '#751b1f' },
              }}
            />
          </Stack>
        );
      },
    },
  ];

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Text variant="xLarge" block style={{ fontWeight: 600 }}>
        Admin Asset Request Approvals
      </Text>

      {message && (
        <MessageBar
          messageBarType={message.type}
          isMultiline={true}
          onDismiss={() => setMessage(null)}
          dismissButtonAriaLabel="Close"
        >
          {message.text}
        </MessageBar>
      )}

      <Card>
        <Card.Section tokens={cardTokens}>
          <Stack tokens={{ childrenGap: 15 }}>
            <SearchBox
              placeholder="Search pending requests..."
              value={searchText}
              onChange={(ev, newValue) => setSearchText(newValue || '')}
            />

            <Text variant="small" style={{ color: '#666' }}>
              Showing {filteredRequests.length} pending request(s)
            </Text>

            {filteredRequests.length > 0 ? (
              <DetailsList
                items={filteredRequests}
                columns={columns}
                setKey="set-items"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            ) : (
              <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
                <Icon iconName="Accept" style={{ fontSize: '48px', color: '#ccc', marginBottom: '10px' }} />
                <Text variant="large" style={{ color: '#666' }}>
                  No pending asset requests found!
                </Text>
              </Stack>
            )}
          </Stack>
        </Card.Section>
      </Card>
    </Stack>
  );
};
