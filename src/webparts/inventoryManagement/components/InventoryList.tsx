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
import styles from './InventoryManagement.module.scss';

export interface IInventoryListProps {
  items: IInventoryItem[];
  isAdmin?: boolean;
  onReturnAsset?: (item: IInventoryItem) => void;
  enablePagination?: boolean;
  pageSize?: number;
}

export const InventoryList: React.FC<IInventoryListProps> = (props) => {
  const [currentPage, setCurrentPage] = React.useState<number>(1);
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
    {
      key: 'columnActivation',
      name: 'Activation State',
      minWidth: 100,
      maxWidth: 120,
      isResizable: true,
      onRender: (item: IInventoryItem) => {
        const statusVal = (item.status || '').toLowerCase();
        let activationState = 'Deactivated';
        let badgeColor = '#991b1b'; // Red
        let bgColor = '#fee2e2';

        if (statusVal === 'assigned') {
          activationState = 'Activated';
          badgeColor = '#166534'; // Green
          bgColor = '#dcfce7';
        } else if (statusVal === 'in stock' || statusVal === 'yes') {
          activationState = 'Inactivated';
          badgeColor = '#92400e'; // Dark orange/amber
          bgColor = '#fef3c7'; // Amber 100
        }

        return (
          <span style={{ 
            backgroundColor: bgColor, 
            color: badgeColor, 
            padding: '4px 12px', 
            borderRadius: '9999px', 
            fontSize: '0.75rem', 
            fontWeight: 600,
            display: 'inline-block'
          }}>
            {activationState}
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
  ];  // Reset page when items array changes
  React.useEffect(() => {
    setCurrentPage(1);
  }, [props.items]);

  const normalizeGroupTitle = (title: string | undefined): string => {
    const t = (title || 'Uncategorized').trim();
    if (/^company\s*assets?$/i.test(t)) return 'Company Assets';
    if (/^leased\s*assets?$/i.test(t)) return 'Leased Assets';
    return t;
  };

  // 1. Process items (sorting by group title if admin)
  let itemsToProcess = props.items;
  if (props.isAdmin && itemsToProcess.length > 0) {
    itemsToProcess = [...props.items].sort((a, b) => 
      normalizeGroupTitle(a.title).localeCompare(normalizeGroupTitle(b.title))
    );
  }

  // 2. Paginate if enabled
  const pageSize = props.pageSize || 10;
  const totalItems = itemsToProcess.length;
  const totalPages = Math.ceil(totalItems / pageSize);
  const activePage = Math.min(currentPage, Math.max(1, totalPages));
  const startIndex = props.enablePagination ? (activePage - 1) * pageSize : 0;
  
  const itemsToRender = props.enablePagination 
    ? itemsToProcess.slice(startIndex, startIndex + pageSize)
    : itemsToProcess;

  // 3. Compute groups based on the actual items to render
  let groups: IGroup[] | undefined = undefined;
  if (props.isAdmin && itemsToRender.length > 0) {
    groups = [];
    let currentGroupName = normalizeGroupTitle(itemsToRender[0].title);
    let currentGroupStartIndex = 0;

    itemsToRender.forEach((item, index) => {
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

    if (itemsToRender.length > 0) {
      groups!.push({
        key: currentGroupName,
        name: currentGroupName,
        startIndex: currentGroupStartIndex,
        count: itemsToRender.length - currentGroupStartIndex,
        isCollapsed: false,
      });
    }
  }

  const getPageNumbers = () => {
    const pages: (number | string)[] = [];
    const maxVisiblePages = 5;
    
    if (totalPages <= maxVisiblePages) {
      for (let i = 1; i <= totalPages; i++) {
        pages.push(i);
      }
    } else {
      pages.push(1);
      
      const start = Math.max(2, activePage - 1);
      const end = Math.min(totalPages - 1, activePage + 1);
      
      if (start > 2) {
        pages.push('...');
      }
      
      for (let i = start; i <= end; i++) {
        pages.push(i);
      }
      
      if (end < totalPages - 1) {
        pages.push('...');
      }
      
      pages.push(totalPages);
    }
    return pages;
  };

  return (
    <div style={{ marginTop: '10px' }}>
      <DetailsList
        items={itemsToRender}
        columns={columns}
        groups={groups}
        setKey="set"
        layoutMode={DetailsListLayoutMode.justified}
        selectionMode={SelectionMode.none}
      />

      {props.enablePagination && totalPages > 1 && (
        <div className={styles.paginationContainer}>
          <div className={styles.paginationInfo}>
            Showing <strong>{startIndex + 1}</strong> to <strong>{Math.min(startIndex + pageSize, totalItems)}</strong> of <strong>{totalItems}</strong> entries
          </div>
          <div className={styles.paginationControls}>
            <button
              className={styles.paginationButton}
              disabled={activePage === 1}
              onClick={() => setCurrentPage(1)}
              title="First Page"
            >
              &laquo;
            </button>
            <button
              className={styles.paginationButton}
              disabled={activePage === 1}
              onClick={() => setCurrentPage(prev => prev - 1)}
              title="Previous Page"
            >
              &lsaquo;
            </button>

            {getPageNumbers().map((page, idx) => {
              if (page === '...') {
                return <span key={`ellipsis-${idx}`} style={{ padding: '0 8px', color: 'var(--text-muted)' }}>...</span>;
              }
              return (
                <button
                  key={page}
                  className={`${styles.paginationButton} ${activePage === page ? styles.active : ''}`}
                  onClick={() => setCurrentPage(page as number)}
                >
                  {page}
                </button>
              );
            })}

            <button
              className={styles.paginationButton}
              disabled={activePage === totalPages}
              onClick={() => setCurrentPage(prev => prev + 1)}
              title="Next Page"
            >
              &rsaquo;
            </button>
            <button
              className={styles.paginationButton}
              disabled={activePage === totalPages}
              onClick={() => setCurrentPage(totalPages)}
              title="Last Page"
            >
              &raquo;
            </button>
          </div>
        </div>
      )}
    </div>
  );
};
