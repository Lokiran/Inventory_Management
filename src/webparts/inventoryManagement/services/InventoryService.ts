import { IInventoryItem } from "../models/IInventoryItem";
import { IRequest } from "../models/IRequest";
import { IEventLog } from "../models/IEventLog";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";

export class InventoryService {
  private static readonly LIST_NAME = "InventoryList";
  private static readonly EVENT_LOG_LIST = "EventLogList";

  public static async getItems(sp: SPFI): Promise<IInventoryItem[]> {
    try {
      const items = await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items();

      return items.map(item => ({
        id: item.ID.toString(),
        title: item.Title || "",
        assetName: item.AssetName || "",
        assetType: item.AssetType || "",
        serialNumber: item.SerialNumber || "",
        purchaseDate: item.PurchaseDate || "",
        status: item.Status || "",
        assignedTo: item.AssignedTo ? item.AssignedTo.Title : ""
      }));
    } catch (error: any) {
      console.error("Error fetching items from SharePoint:", error);
      throw error;
    }
  }

  public static async addItem(sp: SPFI, item: Omit<IInventoryItem, 'id'>, userDisplayName: string = "Unknown"): Promise<void> {
    try {
      const addedItem = await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items.add({
        Title: item.title,
        AssetName: item.assetName,
        AssetType: item.assetType,
        SerialNumber: item.serialNumber,
        PurchaseDate: item.purchaseDate,
        Status: item.status
      });

      // Log the event
      await this.addAuditLog(sp, {
        title: `Created Asset: ${item.title}`,
        action: 'Create',
        entityType: 'Asset',
        entityId: addedItem.data.Id ? addedItem.data.Id.toString() : 'Unknown',
        details: JSON.stringify(item),
        user: userDisplayName
      });
    } catch (error: any) {
      console.error("Error adding item to SharePoint:", error);
      throw error;
    }
  }

  public static async addRequest(sp: SPFI, request: Omit<IRequest, 'id' | 'status' | 'requestDate'>, userDisplayName: string = "Unknown"): Promise<void> {
    const payloads = [
      { // Guess 1: Spaces removed
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        SelectAsset: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
      },
      { // Guess 2: Spaces replaced with _x0020_
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        Select_x0020_Asset: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
      },
      { // Guess 3: Title renamed to Employee
        Title: request.requesterName,
        SelectAsset: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
      },
      { // Guess 4: Title renamed to Employee, spaces replaced
        Title: request.requesterName,
        Select_x0020_Asset: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
      },
      { // Guess 5: Select Asset is a Lookup column
        Title: request.requesterName,
        SelectAssetId: parseInt(request.assetId),
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
      },
      { // Guess 6: Select Asset is a Lookup column, Employee is a Person column (might fail if requesterName is string, but worth a shot)
        Title: `Request for ${request.assetTitle}`,
        EmployeeId: parseInt(request.requesterName) || 1, // Will fail gracefully if not a number
        SelectAssetId: parseInt(request.assetId),
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
      }
    ];

    let lastError: any;
    for (const payload of payloads) {
      try {
        const addedRequest = await sp.web.lists.getByTitle("RequestList").items.add(payload);
        
        // Log the event
        await this.addAuditLog(sp, {
          title: `Created Request for Asset: ${request.assetTitle}`,
          action: 'Create',
          entityType: 'Request',
          entityId: addedRequest.data.Id ? addedRequest.data.Id.toString() : 'Unknown',
          details: JSON.stringify(request),
          user: userDisplayName
        });
        
        return; // Success!
      } catch (err) {
        lastError = err;
        // Keep trying the next payload
      }
    }

    console.error("Error adding request to SharePoint after trying all column combinations.", lastError);
    throw new Error(`SharePoint rejected the save. The columns you created in RequestList do not match the expected format. Please check the Developer Console (F12) for the exact column name mismatch.`);
  }

  public static async deleteItem(sp: SPFI, id: number, itemTitle: string = "Unknown", userDisplayName: string = "Unknown"): Promise<void> {
    try {
      await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items.getById(id).delete();
      
      // Log the event
      await this.addAuditLog(sp, {
        title: `Deleted Asset: ${itemTitle}`,
        action: 'Delete',
        entityType: 'Asset',
        entityId: id.toString(),
        details: `Deleted asset with ID ${id}`,
        user: userDisplayName
      });
    } catch (error: any) {
      console.error("Error deleting item from SharePoint:", error);
      throw error;
    }
  }

  public static async getRequests(sp: SPFI): Promise<IRequest[]> {
    try {
      const items = await sp.web.lists.getByTitle("RequestList").items();

      return items.map((item: any) => {
        const keys = Object.keys(item);
        const findKey = (searchStr: string) => keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').includes(searchStr));

        const employeeKey = findKey("employee") || "Employee";
        const selectAssetKey = findKey("selectasset") || "SelectAsset";
        const quantityKey = findKey("quantity") || "Quantity";
        const reasonKey = findKey("reason") || "ReasonforRequest";

        return {
          id: item.ID ? item.ID.toString() : Math.random().toString(36).substr(2, 9),
          requesterName: item[employeeKey] || item.Title || "",
          assetId: "",
          assetTitle: item[selectAssetKey] || item.Title || "",
          quantity: parseInt(item[quantityKey]) || 1,
          status: item.Status || 'Pending',
          requestDate: item.Created ? item.Created.split('T')[0] : new Date().toISOString().split('T')[0],
          reason: item[reasonKey] || ""
        };
      });
    } catch (error: any) {
      console.error("Error fetching requests from SharePoint:", error);
      throw error;
    }
  }

  public static async addAuditLog(sp: SPFI, log: Omit<IEventLog, 'id' | 'timestamp'>): Promise<void> {
    try {
      await sp.web.lists.getByTitle(InventoryService.EVENT_LOG_LIST).items.add({
        Title: log.title,
        Action: log.action,
        EntityType: log.entityType,
        EntityId: log.entityId,
        Details: log.details,
        User: log.user
      });
    } catch (error: any) {
      console.error("Error adding audit log to SharePoint:", error);
      // We don't throw here to avoid failing the main operation if logging fails
    }
  }

  public static async getAuditLogs(sp: SPFI): Promise<IEventLog[]> {
    let logs: IEventLog[] = [];
    try {
      // 1. Fetch from InventoryList
      try {
        const inventoryItems = await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items.select("ID", "Title", "AssetName", "Created", "Modified", "Author/Title", "Editor/Title").expand("Author", "Editor")();
        
        inventoryItems.forEach(item => {
           logs.push({
             id: `asset-create-${item.ID}`,
             title: `Created Asset: ${item.Title}`,
             action: 'Create',
             entityType: 'Asset',
             entityId: item.ID.toString(),
             assetName: item.AssetName || item.Title,
             details: `Asset was added to inventory`,
             user: item.Author?.Title || "System",
             timestamp: item.Created ? item.Created.split('T')[0] + ' ' + item.Created.split('T')[1].substring(0, 8) : new Date().toISOString()
           });
           
           if (item.Modified && item.Created && new Date(item.Modified).getTime() - new Date(item.Created).getTime() > 5000) {
             logs.push({
               id: `asset-update-${item.ID}-${new Date(item.Modified).getTime()}`,
               title: `Updated Asset: ${item.Title}`,
               action: 'Update',
               entityType: 'Asset',
               entityId: item.ID.toString(),
               assetName: item.AssetName || item.Title,
               details: `Asset details were modified`,
               user: item.Editor?.Title || "System",
               timestamp: item.Modified.split('T')[0] + ' ' + item.Modified.split('T')[1].substring(0, 8)
             });
           }
        });
      } catch (err) {
        console.warn("Could not fetch InventoryList for audit logs", err);
      }

      // 2. Fetch from RequestList
      try {
        const requestItems = await sp.web.lists.getByTitle("RequestList").items();
        
        requestItems.forEach((item: any) => {
           const keys = Object.keys(item);
           const findKey = (searchStr: string) => keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').includes(searchStr));

           const employeeKey = findKey("employee") || "Employee";
           const selectAssetKey = findKey("selectasset") || "SelectAsset";

           const reqAssetName = item[selectAssetKey] || item.Title || "Unknown Asset";
           const reqUser = item[employeeKey] || item.Title || "System";

           logs.push({
             id: `request-create-${item.ID}`,
             title: `Created Request: ${reqAssetName}`,
             action: 'Create',
             entityType: 'Request',
             entityId: item.ID.toString(),
             assetName: reqAssetName,
             details: `Asset request was submitted`,
             user: reqUser,
             timestamp: item.Created ? item.Created.split('T')[0] + ' ' + item.Created.split('T')[1].substring(0, 8) : new Date().toISOString()
           });
           
           if (item.Modified && item.Created && new Date(item.Modified).getTime() - new Date(item.Created).getTime() > 5000) {
             logs.push({
               id: `request-update-${item.ID}-${new Date(item.Modified).getTime()}`,
               title: `Updated Request: ${reqAssetName}`,
               action: 'Update',
               entityType: 'Request',
               entityId: item.ID.toString(),
               assetName: reqAssetName,
               details: `Request details were modified`,
               user: reqUser,
               timestamp: item.Modified.split('T')[0] + ' ' + item.Modified.split('T')[1].substring(0, 8)
             });
           }
        });
      } catch (err) {
        console.warn("Could not fetch RequestList for audit logs", err);
      }

      // Sort logs by timestamp descending
      logs.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

      return logs;
    } catch (error: any) {
      console.error("Error generating audit logs from SharePoint lists:", error);
      return []; 
    }
  }
}
