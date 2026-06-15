import * as React from 'react';
import { IInventoryItem } from '../models/IInventoryItem';
import { 
  DetailsList, 
  DetailsListLayoutMode, 
  SelectionMode, 
  IColumn,
  IGroup
} from '@fluentui/react/lib/DetailsList';
import { PrimaryButton } from '@fluentui/react/lib/Button';

export interface IInventoryListProps {
  items: IInventoryItem[];
  isAdmin?: boolean;
  onReturnAsset?: (item: IInventoryItem) => void;
}

export const InventoryList: React.FC<IInventoryListProps> = (props) => {
  const columns: IColumn[] = [
    { key: 'column1', name: 'ID', fieldName: 'id', minWidth: 40, maxWidth: 40, isResizable: true },
    { key: 'column2', name: 'Title', fieldName: 'title', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column3', name: 'Asset Name', fieldName: 'assetName', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column4', name: 'Type', fieldName: 'assetType', minWidth: 80, maxWidth: 100, isResizable: true },
    { key: 'column5', name: 'Serial Number', fieldName: 'serialNumber', minWidth: 100, maxWidth: 120, isResizable: true },
    { key: 'column6', name: 'Purchase Date', fieldName: 'purchaseDate', minWidth: 100, maxWidth: 120, isResizable: true },
    { key: 'columnVendor', name: 'Vendor', fieldName: 'vendor', minWidth: 80, maxWidth: 100, isResizable: true },
    { key: 'columnCondition', name: 'Condition', fieldName: 'condition', minWidth: 80, maxWidth: 100, isResizable: true },
    { key: 'columnWarranty', name: 'Warranty Expiry', fieldName: 'warrantyExpiry', minWidth: 100, maxWidth: 120, isResizable: true },
    { 
      key: 'column7', 
      name: 'Status', 
      fieldName: 'status', 
      minWidth: 80, 
      maxWidth: 100, 
      isResizable: true,
      onRender: (item: IInventoryItem) => {
        const isAvailable = item.status === 'Yes' || item.status === 'In Stock';
        const isPendingReturn = item.status === 'Pending Return';
        const isReturnApproved = item.status === 'Return Approved';
        
        let backgroundColor = '#fee2e2';
        let textColor = '#991b1b';

        if (isAvailable) {
          backgroundColor = '#dcfce7';
          textColor = '#166534';
        } else if (isPendingReturn) {
          backgroundColor = '#ffedd5';
          textColor = '#9a3412';
        } else if (isReturnApproved) {
          backgroundColor = '#dcfce7';
          textColor = '#166534';
        }
        
        return (
          <span style={{ 
            backgroundColor, 
            color: textColor, 
            padding: '4px 12px', 
            borderRadius: '9999px', 
            fontSize: '0.75rem', 
            fontWeight: 600,
            display: 'inline-block'
          }}>
            {item.status}
          </span>
        );
      }
    },
    { key: 'column8', name: 'Assigned To', fieldName: 'assignedTo', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column9', name: 'Specifications', fieldName: 'specifications', minWidth: 150, maxWidth: 300, isResizable: true },
    ...(props.onReturnAsset ? [
      {
        key: 'columnActions',
        name: 'Actions',
        minWidth: 100,
        maxWidth: 120,
        isResizable: true,
        onRender: (item: IInventoryItem) => {
          const isPendingReturn = item.status === 'Pending Return';
          const isReturnApproved = item.status === 'Return Approved';
          
          if (isPendingReturn) {
            return (
              <span style={{ color: '#ea580c', fontWeight: 600, fontSize: '0.8rem' }}>
                Pending Return
              </span>
            );
          }
          if (isReturnApproved) {
            return (
              <span style={{ color: '#16a34a', fontWeight: 600, fontSize: '0.8rem' }}>
                Approved
              </span>
            );
          }
          return (
            <PrimaryButton
              text="Return"
              onClick={() => props.onReturnAsset!(item)}
              styles={{ root: { height: 26, padding: '4px 8px', fontSize: '0.8rem' } }}
            />
          );
        }
      }
    ] : [])
  ];

  // Grouping Logic for Admins
  let items = props.items;
  let groups: IGroup[] | undefined = undefined;

  if (props.isAdmin && items.length > 0) {
    const normalizeGroupTitle = (title: string | undefined): string => {
      const t = (title || 'Uncategorized').trim();
      if (/^company\s*assets?$/i.test(t)) return 'Company Assets';
      if (/^leased\s*assets?$/i.test(t)) return 'Leased Assets';
      return t;
    };

    // Sort items by normalized title to ensure correct grouping
    items = [...props.items].sort((a, b) => 
      normalizeGroupTitle(a.title).localeCompare(normalizeGroupTitle(b.title))
    );
    
    groups = [];
    let currentGroupName = normalizeGroupTitle(items[0].title);
    let currentGroupStartIndex = 0;

    items.forEach((item, index) => {
      const itemGroup = normalizeGroupTitle(item.title);
      if (itemGroup !== currentGroupName) {
        groups!.push({
          key: currentGroupName,
          name: currentGroupName,
          startIndex: currentGroupStartIndex,
          count: index - currentGroupStartIndex,
          isCollapsed: false,
        });
        currentGroupName = itemGroup;
        currentGroupStartIndex = index;
      }
    });

    // Add the last group
    if (items.length > 0) {
      groups!.push({
        key: currentGroupName,
        name: currentGroupName,
        startIndex: currentGroupStartIndex,
        count: items.length - currentGroupStartIndex,
        isCollapsed: false,
      });
    }
  }

  return (
    <div style={{ marginTop: '10px' }}>
      <DetailsList
        items={items}
        columns={columns}
        groups={groups}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
      />
    </div>
  );
};
