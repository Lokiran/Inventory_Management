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
import styles from './MyRequests.module.scss';
import { IEmployeeManagementProps } from '../IEmployeeManagementProps';
import { InventoryService } from '../../services/InventoryService';

interface IAssetRequestItem {
  id: string;
  assetType: string;
  assetName: string;
  quantity: number;
  priority: string;
  status: string;
  requestDate: string;
  requiredDate: string;
  description: string;
}

const cardTokens: ICardTokens = { childrenMargin: 12 };

const statusBadgeStyles: { [key: string]: { backgroundColor: string; color: string } } = {
  Pending: { backgroundColor: '#ffb81c', color: '#fff' },
  Approved: { backgroundColor: '#27ae60', color: '#fff' },
  Rejected: { backgroundColor: '#e74c3c', color: '#fff' },
  Issued: { backgroundColor: '#0078d4', color: '#fff' },
};

export const MyRequests: React.FC<IEmployeeManagementProps & { setIsLoading: (loading: boolean) => void }> = (props) => {
  const [requests, setRequests] = useState<IAssetRequestItem[]>([]);
  const [filteredRequests, setFilteredRequests] = useState<IAssetRequestItem[]>([]);
  const [searchText, setSearchText] = useState('');
  const [statusFilter, setStatusFilter] = useState<string | null>(null);
  const [selectedRequest, setSelectedRequest] = useState<IAssetRequestItem | null>(null);
  const [showDetailDialog, setShowDetailDialog] = useState(false);

  useEffect(() => {
    loadRequests();
  }, []);

  useEffect(() => {
    filterRequests();
  }, [searchText, statusFilter, requests]);

  const loadRequests = async () => {
    try {
      props.setIsLoading(true);
      const service = new InventoryService(props.apiBaseUrl);
      const data = await service.getEmployeeAssetRequests(props.userEmail);
      setRequests(data);
    } catch (error) {
      console.error('Error loading requests:', error);
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
          req.assetType.toLowerCase().includes(searchText.toLowerCase())
      );
    }

    if (statusFilter) {
      filtered = filtered.filter((req) => req.status === statusFilter);
    }

    setFilteredRequests(filtered);
  };

  const handleViewDetails = (item: IAssetRequestItem) => {
    setSelectedRequest(item);
    setShowDetailDialog(true);
  };

  const handleCancelRequest = async (requestId: string) => {
    try {
      const service = new InventoryService(props.apiBaseUrl);
      await service.cancelAssetRequest(requestId);
      loadRequests();
    } catch (error) {
      console.error('Error canceling request:', error);
    }
  };

  const columns: IColumn[] = [
    {
      key: 'assetType',
      name: 'Asset Type',
      minWidth: 100,
      onRender: (item: IAssetRequestItem) => <Text>{item.assetType}</Text>,
    },
    {
      key: 'assetName',
      name: 'Asset Name',
      minWidth: 150,
      onRender: (item: IAssetRequestItem) => <Text>{item.assetName}</Text>,
    },
    {
      key: 'quantity',
      name: 'Quantity',
      minWidth: 80,
      onRender: (item: IAssetRequestItem) => <Text>{item.quantity}</Text>,
    },
    {
      key: 'priority',
      name: 'Priority',
      minWidth: 100,
      onRender: (item: IAssetRequestItem) => (
        <span style={{ padding: '4px 8px', borderRadius: '4px', backgroundColor: item.priority === 'high' ? '#d13438' : item.priority === 'urgent' ? '#a4262c' : '#0078d4', color: 'white', fontSize: '12px', fontWeight: 600 }}>
          {item.priority.charAt(0).toUpperCase() + item.priority.slice(1)}
        </span>
      ),
    },
    {
      key: 'status',
      name: 'Status',
      minWidth: 120,
      onRender: (item: IAssetRequestItem) => (
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
      key: 'requestDate',
      name: 'Request Date',
      minWidth: 120,
      onRender: (item: IAssetRequestItem) => (
        <Text>{new Date(item.requestDate).toLocaleDateString()}</Text>
      ),
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 150,
      onRender: (item: IAssetRequestItem) => (
        <Stack horizontal tokens={{ childrenGap: 10 }}>
          <PrimaryButton
            text="View"
            onClick={() => handleViewDetails(item)}
            styles={{
              root: { padding: '4px 12px', fontSize: '12px', height: '24px' },
            }}
          />
          {item.status === 'Pending' && (
            <PrimaryButton
              text="Cancel"
              onClick={() => handleCancelRequest(item.id)}
              styles={{
                root: { padding: '4px 12px', fontSize: '12px', height: '24px', backgroundColor: '#e74c3c' },
              }}
            />
          )}
        </Stack>
      ),
    },
  ];

  const statusFilterOptions: IDropdownOption[] = [
    { key: '', text: 'All Status' },
    { key: 'Pending', text: 'Pending' },
    { key: 'Approved', text: 'Approved' },
    { key: 'Rejected', text: 'Rejected' },
    { key: 'Issued', text: 'Issued' },
  ];

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Text variant="xLarge" block style={{ fontWeight: 600 }}>
        My Asset Requests
      </Text>

      <Card>
        <Card.Section tokens={cardTokens}>
          <Stack tokens={{ childrenGap: 15 }}>
            {/* Filters */}
            <Stack horizontal tokens={{ childrenGap: 15 }}>
              <SearchBox
                placeholder="Search by asset name or type..."
                value={searchText}
                onChange={(ev, newValue) => setSearchText(newValue || '')}
                style={{ flex: 1 }}
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
              Showing {filteredRequests.length} of {requests.length} requests
            </Text>

            {/* Details List */}
            {filteredRequests.length > 0 ? (
              <DetailsList
                items={filteredRequests}
                columns={columns}
                setKey="set-items"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            ) : (
              <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '300px' }}>
                <Icon iconName="ClearFilter" style={{ fontSize: '48px', color: '#ccc', marginBottom: '10px' }} />
                <Text variant="large" style={{ color: '#666' }}>
                  No requests found.
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
          title: 'Request Details',
          closeButtonAriaLabel: 'Close',
        }}
      >
        {selectedRequest && (
          <Stack tokens={{ childrenGap: 15 }}>
            <TextField label="Asset Type" value={selectedRequest.assetType} disabled />
            <TextField label="Asset Name" value={selectedRequest.assetName} disabled />
            <TextField label="Quantity" value={selectedRequest.quantity.toString()} disabled />
            <TextField label="Priority" value={selectedRequest.priority} disabled />
            <TextField label="Status" value={selectedRequest.status} disabled />
            <TextField label="Description" value={selectedRequest.description} multiline rows={4} disabled />
            <TextField
              label="Required Date"
              value={new Date(selectedRequest.requiredDate).toLocaleDateString()}
              disabled
            />
          </Stack>
        )}
        <DialogFooter>
          <PrimaryButton text="Close" onClick={() => setShowDetailDialog(false)} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

