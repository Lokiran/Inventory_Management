import * as React from 'react';
import { IRequest } from '../models/IRequest';
import { 
  DetailsList, 
  DetailsListLayoutMode, 
  SelectionMode, 
  IColumn 
} from '@fluentui/react/lib/DetailsList';

export interface IRequestListProps {
  items: IRequest[];
}

export const RequestList: React.FC<IRequestListProps> = (props) => {
  const columns: IColumn[] = [
    { key: 'column1', name: 'ID', fieldName: 'id', minWidth: 60, maxWidth: 80, isResizable: true },
    { key: 'column2', name: 'Requester', fieldName: 'requesterName', minWidth: 100, maxWidth: 150, isResizable: true },
    { key: 'column3', name: 'Asset Title', fieldName: 'assetTitle', minWidth: 120, maxWidth: 200, isResizable: true },
    { key: 'column4', name: 'Quantity', fieldName: 'quantity', minWidth: 60, maxWidth: 80, isResizable: true },
    { key: 'column5', name: 'Date', fieldName: 'requestDate', minWidth: 100, maxWidth: 120, isResizable: true },
    { 
      key: 'column6', 
      name: 'Status', 
      fieldName: 'status', 
      minWidth: 80, 
      maxWidth: 100, 
      isResizable: true,
      onRender: (item: IRequest) => {
        let backgroundColor = '#fef3c7'; // default pending (yellow)
        let textColor = '#92400e';
        
        if (item.status === 'Approved') {
          backgroundColor = '#dcfce7';
          textColor = '#166534';
        } else if (item.status === 'Declined') {
          backgroundColor = '#fee2e2';
          textColor = '#991b1b';
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
    { key: 'column7', name: 'Reason', fieldName: 'reason', minWidth: 150, maxWidth: 250, isResizable: true },
  ];

  return (
    <div style={{ marginTop: '10px' }}>
      {props.items.length === 0 ? (
        <p style={{ fontStyle: 'italic', color: 'var(--text-muted)' }}>No asset requests found.</p>
      ) : (
        <DetailsList
          items={props.items}
          columns={columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionMode={SelectionMode.none}
        />
      )}
    </div>
  );
};
