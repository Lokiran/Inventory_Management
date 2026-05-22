import * as React from 'react';
import { IInventoryItem } from '../models/IInventoryItem';
import { 
  DetailsList, 
  DetailsListLayoutMode, 
  SelectionMode, 
  IColumn,
  IGroup
} from '@fluentui/react/lib/DetailsList';

export interface IInventoryListProps {
  items: IInventoryItem[];
  isAdmin?: boolean;
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
    { key: 'column9', name: 'Note', fieldName: 'note', minWidth: 150, maxWidth: 300, isResizable: true },
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
