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
  Dialog,
  DialogType,
  DialogFooter,
  TextField,
} from '@fluentui/react';
import { Card, ICardTokens } from '@uifabric/react-cards';
import styles from './MyAssignedAssets.module.scss';
import { IEmployeeManagementProps } from '../IEmployeeManagementProps';
import { InventoryService } from '../../services/InventoryService';

interface IAssignedAsset {
  id: string;
  assetType: string;
  assetName: string;
  serialNumber: string;
  assignmentDate: string;
  status: string;
  condition: string;
  location: string;
}

const cardTokens: ICardTokens = { childrenMargin: 12 };

export const MyAssignedAssets: React.FC<IEmployeeManagementProps & { setIsLoading: (loading: boolean) => void }> = (props) => {
  const [assets, setAssets] = useState<IAssignedAsset[]>([]);
  const [filteredAssets, setFilteredAssets] = useState<IAssignedAsset[]>([]);
  const [searchText, setSearchText] = useState('');
  const [selectedAsset, setSelectedAsset] = useState<IAssignedAsset | null>(null);
  const [showDetailDialog, setShowDetailDialog] = useState(false);

  useEffect(() => {
    loadAssets();
  }, []);

  useEffect(() => {
    filterAssets();
  }, [searchText, assets]);

  const loadAssets = async () => {
    try {
      props.setIsLoading(true);
      const service = new InventoryService(props.apiBaseUrl);
      const data = await service.getEmployeeAssignedAssets(props.userEmail);
      setAssets(data);
    } catch (error) {
      console.error('Error loading assigned assets:', error);
    } finally {
      props.setIsLoading(false);
    }
  };

  const filterAssets = () => {
    let filtered = [...assets];

    if (searchText) {
      filtered = filtered.filter(
        (asset) =>
          asset.assetName.toLowerCase().includes(searchText.toLowerCase()) ||
          asset.serialNumber.toLowerCase().includes(searchText.toLowerCase()) ||
          asset.assetType.toLowerCase().includes(searchText.toLowerCase())
      );
    }

    setFilteredAssets(filtered);
  };

  const handleViewDetails = (item: IAssignedAsset) => {
    setSelectedAsset(item);
    setShowDetailDialog(true);
  };

  const handleDownloadAssetInfo = (asset: IAssignedAsset) => {
    const assetInfo = `
Asset Information Report
==========================
Asset Type: ${asset.assetType}
Asset Name: ${asset.assetName}
Serial Number: ${asset.serialNumber}
Assignment Date: ${new Date(asset.assignmentDate).toLocaleDateString()}
Status: ${asset.status}
Condition: ${asset.condition}
Location: ${asset.location}
Generated: ${new Date().toLocaleString()}
    `;
    const element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(assetInfo));
    element.setAttribute('download', `asset-${asset.serialNumber}.txt`);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
  };

  const columns: IColumn[] = [
    {
      key: 'assetType',
      name: 'Type',
      minWidth: 100,
      onRender: (item: IAssignedAsset) => <Text>{item.assetType}</Text>,
    },
    {
      key: 'assetName',
      name: 'Asset Name',
      minWidth: 150,
      onRender: (item: IAssignedAsset) => <Text>{item.assetName}</Text>,
    },
    {
      key: 'serialNumber',
      name: 'Serial Number',
      minWidth: 120,
      onRender: (item: IAssignedAsset) => <Text>{item.serialNumber}</Text>,
    },
    {
      key: 'condition',
      name: 'Condition',
      minWidth: 100,
      onRender: (item: IAssignedAsset) => (
        <Text style={{ color: item.condition === 'Good' ? '#27ae60' : item.condition === 'Fair' ? '#f39c12' : '#e74c3c' }}>
          {item.condition}
        </Text>
      ),
    },
    {
      key: 'assignmentDate',
      name: 'Assigned Date',
      minWidth: 120,
      onRender: (item: IAssignedAsset) => (
        <Text>{new Date(item.assignmentDate).toLocaleDateString()}</Text>
      ),
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 150,
      onRender: (item: IAssignedAsset) => (
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
            onClick={() => handleDownloadAssetInfo(item)}
            styles={{
              root: { padding: '4px 12px', fontSize: '12px', height: '24px' },
            }}
          />
        </Stack>
      ),
    },
  ];

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Text variant="xLarge" block style={{ fontWeight: 600 }}>
        My Assigned Assets
      </Text>

      <Card>
        <Card.Section tokens={cardTokens}>
          <Stack tokens={{ childrenGap: 15 }}>
            {/* Search */}
            <SearchBox
              placeholder="Search by asset name, type, or serial number..."
              value={searchText}
              onChange={(ev, newValue) => setSearchText(newValue || '')}
            />

            {/* Items Count */}
            <Text variant="small" style={{ color: '#666' }}>
              You have {filteredAssets.length} assigned asset(s)
            </Text>

            {/* Details List */}
            {filteredAssets.length > 0 ? (
              <DetailsList
                items={filteredAssets}
                columns={columns}
                setKey="set-items"
                layoutMode={DetailsListLayoutMode.justified}
                selectionMode={SelectionMode.none}
              />
            ) : (
              <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '300px' }}>
                <Icon iconName="BackToWindow" style={{ fontSize: '48px', color: '#ccc', marginBottom: '10px' }} />
                <Text variant="large" style={{ color: '#666' }}>
                  No assigned assets found.
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
          title: 'Asset Details',
          closeButtonAriaLabel: 'Close',
        }}
      >
        {selectedAsset && (
          <Stack tokens={{ childrenGap: 15 }}>
            <TextField label="Asset Type" value={selectedAsset.assetType} disabled />
            <TextField label="Asset Name" value={selectedAsset.assetName} disabled />
            <TextField label="Serial Number" value={selectedAsset.serialNumber} disabled />
            <TextField label="Condition" value={selectedAsset.condition} disabled />
            <TextField label="Location" value={selectedAsset.location} disabled />
            <TextField
              label="Assigned Date"
              value={new Date(selectedAsset.assignmentDate).toLocaleDateString()}
              disabled
            />
            <TextField label="Status" value={selectedAsset.status} disabled />
          </Stack>
        )}
        <DialogFooter>
          <PrimaryButton text="Close" onClick={() => setShowDetailDialog(false)} />
        </DialogFooter>
      </Dialog>
    </Stack>
  );
};

