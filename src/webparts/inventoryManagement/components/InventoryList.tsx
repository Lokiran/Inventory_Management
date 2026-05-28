import * as React from 'react';
import { IInventoryItem } from '../models/IInventoryItem';
import { 
  DetailsList, 
  DetailsListLayoutMode, 
  SelectionMode, 
  IColumn 
} from '@fluentui/react/lib/DetailsList';

export interface IInventoryListProps {
  items: IInventoryItem[];
}

export const InventoryList: React.FC<IInventoryListProps> = (props) => {
  const columns: IColumn[] = [
    { key: 'column1', name: 'ID', fieldName: 'id', minWidth: 40, maxWidth: 40, isResizable: true },
    { key: 'column2', name: 'Title', fieldName: 'title', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column3', name: 'Asset Name', fieldName: 'assetName', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column4', name: 'Type', fieldName: 'assetType', minWidth: 80, maxWidth: 100, isResizable: true },
    { key: 'column5', name: 'Serial Number', fieldName: 'serialNumber', minWidth: 100, maxWidth: 120, isResizable: true },
    { key: 'column6', name: 'Purchase Date', fieldName: 'purchaseDate', minWidth: 100, maxWidth: 120, isResizable: true },
    { 
      key: 'column7', 
      name: 'Status', 
      fieldName: 'status', 
      minWidth: 80, 
      maxWidth: 100, 
      isResizable: true,
      onRender: (item: IInventoryItem) => {
        const isAvailable = item.status === 'Yes' || item.status === 'In Stock';
        const backgroundColor = isAvailable ? '#dcfce7' : '#fee2e2';
        const textColor = isAvailable ? '#166534' : '#991b1b';
        
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
  ];

  return (
    <div style={{ marginTop: '10px' }}>
      <DetailsList
        items={props.items}
        columns={columns}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
      />
    </div>
  );
};
