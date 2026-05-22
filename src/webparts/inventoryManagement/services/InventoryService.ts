import { IInventoryItem } from "../models/IInventoryItem";
import { IRequest } from "../models/IRequest";
import { IEventLog } from "../models/IEventLog";
import { getSP } from "../pnpjsConfig";

export class InventoryService {
  private static readonly LIST_NAME = "InventoryList";
  private static readonly EVENT_LOG_LIST = "EventLogList";
  private static readonly REQUEST_LIST_NAME = "RequestList";
  private static readonly REQUEST_STATUS_INTERNAL_NAME = "RequestStatus";
  private static readonly REQUEST_COMMENT_INTERNAL_NAME = "ManagerComment";
  private static readonly REQUEST_KEY_INTERNAL_NAME = "RequestKey";
  private static readonly ASSET_STATUS_INTERNAL_NAME = "AssetStatus";
  private static _requestWorkflowFieldsEnsured = false;

  private static _normalizeRequestKey(input: string): string {
    return (input || "").trim().toUpperCase();
  }

  private static _buildRequestKeyFromItemId(itemId: number): string {
    const raw = itemId.toString();
    const padded = ("000000" + raw).slice(-6);
    return `REQ-${padded}`;
  }

  private static _extractRequestKey(item: any): string {
    if (!item) {
      return "";
    }

    const directValue = item[InventoryService.REQUEST_KEY_INTERNAL_NAME];
    if (directValue) {
      return this._normalizeRequestKey(directValue.toString());
    }

    const matchingKey = Object.keys(item).find(key =>
      key.toLowerCase().replace(/_x0020_/g, "") === InventoryService.REQUEST_KEY_INTERNAL_NAME.toLowerCase()
    );
    if (matchingKey && item[matchingKey]) {
      return this._normalizeRequestKey(item[matchingKey].toString());
    }

    if (item.ID) {
      return this._buildRequestKeyFromItemId(parseInt(item.ID.toString(), 10));
    }

    return "";
  }

  private static _isBusinessStatusKey(key: string): boolean {
    const lower = (key || '').toLowerCase();
    const normalized = lower.replace(/_x0020_/g, '');
    const isSystemKey = lower.indexOf('__') === 0 || lower.indexOf('odata') >= 0;
    const isModeration = normalized.indexOf('moderationstatus') >= 0 || lower.indexOf('moderation') >= 0;
    const looksLikeStatus = normalized === 'status' || normalized === 'requeststatus' || normalized.indexOf('requeststatus') >= 0;
    return !isSystemKey && !isModeration && looksLikeStatus;
  }

  private static async _ensureRequestWorkflowFields(): Promise<void> {
    const sp = getSP();
    if (this._requestWorkflowFieldsEnsured) {
      return;
    }

    try {
      const list = sp.web.lists.getByTitle(InventoryService.REQUEST_LIST_NAME);
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString")();

      const hasRequestStatus = fields.some(field => {
        const internalName = (field.InternalName || '').toString().toLowerCase();
        return internalName === InventoryService.REQUEST_STATUS_INTERNAL_NAME.toLowerCase();
      });

      const hasManagerComment = fields.some(field => {
        const internalName = (field.InternalName || '').toString().toLowerCase();
        return internalName === InventoryService.REQUEST_COMMENT_INTERNAL_NAME.toLowerCase();
      });

      if (!hasRequestStatus) {
        try {
          await list.fields.addChoice(InventoryService.REQUEST_STATUS_INTERNAL_NAME, {
            Choices: ["Pending", "Approved", "Rejected"],
            FillInChoice: false
          });
        } catch (err) {
          console.warn("Could not auto-create RequestStatus field. Continuing.", err);
        }
      }

      if (!hasManagerComment) {
        try {
          await list.fields.addMultilineText(InventoryService.REQUEST_COMMENT_INTERNAL_NAME);
        } catch (err) {
          console.warn("Could not auto-create ManagerComment field. Continuing.", err);
        }
      }

      const hasRequestKey = fields.some(field => {
        const internalName = (field.InternalName || '').toString().toLowerCase();
        return internalName === InventoryService.REQUEST_KEY_INTERNAL_NAME.toLowerCase();
      });
      if (!hasRequestKey) {
        try {
          await list.fields.addText(InventoryService.REQUEST_KEY_INTERNAL_NAME);
        } catch (err) {
          console.warn("Could not auto-create RequestKey field. Continuing.", err);
        }
      }

      const hasAssetStatus = fields.some(field => {
        const internalName = (field.InternalName || '').toString().toLowerCase();
        return internalName === InventoryService.ASSET_STATUS_INTERNAL_NAME.toLowerCase();
      });
      if (!hasAssetStatus) {
        try {
          await list.fields.addChoice(InventoryService.ASSET_STATUS_INTERNAL_NAME, {
            Choices: ["Pending", "Approved"],
            FillInChoice: false
          });
        } catch (err) {
          console.warn("Could not auto-create AssetStatus field. Continuing.", err);
        }
      }
    } catch (error) {
      // Non-admin users may not have schema permissions. Don't block request flows.
      console.warn("Could not ensure RequestList workflow fields. Continuing with fallback behavior.", error);
    } finally {
      this._requestWorkflowFieldsEnsured = true;
    }
  }

  public static async getItems(): Promise<IInventoryItem[]> {
    const sp = getSP();
    try {
      let items: any[] = [];
      try {
        items = await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items.select("*, AssignedTo/Title").expand("AssignedTo")();
      } catch (expandError) {
        try {
          items = await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items.select("*, Assigned_x0020_To/Title").expand("Assigned_x0020_To")();
        } catch (expandError2) {
          // Fallback for when AssignedTo is just a plain Text column (cannot be expanded)
          items = await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items();
        }
      }

      return items.map((item: any) => ({
        id: item.ID.toString(),
        title: item.Title || "",
        assetName: item.AssetName || item.Asset_x0020_Name || item.Asset || "",
        assetType: item.AssetType || item.Asset_x0020_Type || item.Type || "",
        serialNumber: item.SerialNumber || item.Serial_x0020_Number || "",
        purchaseDate: item.PurchaseDate || item.Purchase_x0020_Date || "",
        status: item.Status || item.AssetStatus || "",
        assignedTo: (() => {
          const assignedField = item.AssignedTo || item.Assigned_x0020_To;
          if (!assignedField) return "";
          if (typeof assignedField === 'string') return assignedField;
          if (Array.isArray(assignedField)) return assignedField.map((a: any) => a.Title).join(', ');
          return assignedField.Title || "";
        })(),
        assignedDate: item.Modified || "",
        warrantyExpiry: item.WarrantyExpiry || item.Warranty_x0020_Expiry || "",
        note: item.Note || item.Notes || ""
      }));
    } catch (error: any) {
      console.error("Error fetching items from SharePoint:", error);
      throw error;
    }
  }

  public static async addItem(item: Omit<IInventoryItem, 'id'>, userDisplayName: string = "Unknown"): Promise<void> {
    const sp = getSP();
    const payloads = [
      {
        Title: item.title,
        AssetName: item.assetName,
        AssetType: item.assetType,
        SerialNumber: item.serialNumber,
        PurchaseDate: item.purchaseDate,
        Status: item.status,
        Note: item.note || ""
      },
      {
        Title: item.title,
        Asset_x0020_Name: item.assetName,
        Asset_x0020_Type: item.assetType,
        Serial_x0020_Number: item.serialNumber,
        Purchase_x0020_Date: item.purchaseDate,
        Status: item.status,
        Note: item.note || ""
      },
      {
        Title: item.title,
        Asset: item.assetName,
        Type: item.assetType,
        Serial_x0020_Number: item.serialNumber,
        Purchase_x0020_Date: item.purchaseDate,
        AssetStatus: item.status,
        Notes: item.note || ""
      }
    ];

    let lastError: any;
    for (const payload of payloads) {
      try {
        const addedItem = await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items.add(payload);
        
        await this.addAuditLog({
          title: `Created Asset: ${item.title}`,
          action: 'Create',
          entityType: 'Asset',
          entityId: addedItem.data.Id ? addedItem.data.Id.toString() : 'Unknown',
          details: JSON.stringify(item),
          user: userDisplayName
        });
        
        return; // Success!
      } catch (error) {
        lastError = error;
      }
    }

    console.error("Error adding item to SharePoint:", lastError);
    throw new Error(`SharePoint rejected the save. The columns you created in InventoryList do not match the expected format. Error: ${lastError.message || JSON.stringify(lastError)}`);
  }

  public static async addRequest(request: Omit<IRequest, 'id' | 'requestKey' | 'status' | 'requestDate'>, userDisplayName: string = "Unknown"): Promise<void> {
    const sp = getSP();
    await this._ensureRequestWorkflowFields();

    const payloads = [
      { // Guess 1: Spaces removed
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        SelectAsset: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 2: Spaces replaced with _x0020_
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        Select_x0020_Asset: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 3: Title renamed to Employee
        Title: request.requesterName,
        SelectAsset: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 4: Title renamed to Employee, spaces replaced
        Title: request.requesterName,
        Select_x0020_Asset: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 5: Select Asset is a Lookup column
        Title: request.requesterName,
        SelectAssetId: parseInt(request.assetId),
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 6: Select Asset is a Lookup column, Employee is a Person column (might fail if requesterName is string, but worth a shot)
        Title: `Request for ${request.assetTitle}`,
        EmployeeId: parseInt(request.requesterName) || 1, // Will fail gracefully if not a number
        SelectAssetId: parseInt(request.assetId),
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 7: Column renamed to AssetType
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        AssetType: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 8: Column renamed to Asset Type (spaces replaced)
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        Asset_x0020_Type: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 9: AssetType as Lookup
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        AssetTypeId: parseInt(request.assetId),
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      },
      { // Guess 10: Fallback - Omit the specific asset column if it's not required
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: "Pending"
      }
    ];

    let lastError: any;
    for (const payload of payloads) {
      try {
        const addedRequest = await sp.web.lists.getByTitle("RequestList").items.add(payload);
        const requestItemId = addedRequest.data.Id ? parseInt(addedRequest.data.Id.toString(), 10) : NaN;
        const requestKey = Number.isNaN(requestItemId)
          ? `REQ-${Date.now().toString(36).toUpperCase()}`
          : this._buildRequestKeyFromItemId(requestItemId);
        if (!Number.isNaN(requestItemId)) {
          try {
            await sp.web.lists
              .getByTitle(InventoryService.REQUEST_LIST_NAME)
              .items.getById(requestItemId)
              .update({
              [InventoryService.REQUEST_KEY_INTERNAL_NAME]: requestKey,
              [InventoryService.ASSET_STATUS_INTERNAL_NAME]: "Pending"
              });
          } catch (err) {
            // If RequestKey column is missing or user can't update it, don't fail request creation.
            console.warn(`Could not persist RequestKey for request item ${requestItemId}.`, err);
          }
        }
        
        // Log the event
        await this.addAuditLog({
          title: `Created Request ${requestKey} for Asset: ${request.assetTitle}`,
          action: 'Create',
          entityType: 'Request',
          entityId: requestKey,
          details: JSON.stringify({
            requestKey,
            lifecycle: "Submitted",
            requesterName: request.requesterName,
            assetTitle: request.assetTitle,
            quantity: request.quantity,
            reason: request.reason || "",
            requestedAt: new Date().toISOString()
          }),
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

  public static async deleteItem(id: number, itemTitle: string = "Unknown", userDisplayName: string = "Unknown"): Promise<void> {
    const sp = getSP();
    try {
      await sp.web.lists.getByTitle(InventoryService.LIST_NAME).items.getById(id).delete();
      
      // Log the event
      await this.addAuditLog({
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

  public static async getRequests(): Promise<IRequest[]> {
    const sp = getSP();
    try {
      await this._ensureRequestWorkflowFields();
      const items = await sp.web.lists.getByTitle(InventoryService.REQUEST_LIST_NAME).items();

      return items.map((item: any) => {
        const keys = Object.keys(item);
        const findKey = (searchStr: string) => keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').includes(searchStr));

        const employeeKey = findKey("employee") || "Employee";
        const selectAssetKey = findKey("selectasset") || "SelectAsset";
        const quantityKey = findKey("quantity") || "Quantity";
        const reasonKey = findKey("reason") || "ReasonforRequest";
        const managerCommentKey = findKey("managercomment") || InventoryService.REQUEST_COMMENT_INTERNAL_NAME;
        const assetStatusKey = findKey("assetstatus") || InventoryService.ASSET_STATUS_INTERNAL_NAME;
        const statusKey = keys.find(key => this._isBusinessStatusKey(key)) || InventoryService.REQUEST_STATUS_INTERNAL_NAME || findKey("status");
        const rawStatus = statusKey ? (item[statusKey] || item.Status || 'Pending') : (item.Status || 'Pending');
        const normalizedStatus = (rawStatus || '').toString().toLowerCase();
        const status =
          (normalizedStatus.includes('approv')) ? 'Approved' :
          (normalizedStatus.includes('declin') || normalizedStatus.includes('reject')) ? 'Declined' :
          'Pending';
        const requestKey = this._extractRequestKey(item);

        return {
          id: item.ID ? item.ID.toString() : Math.random().toString(36).substr(2, 9),
          requestKey: requestKey || (item.ID ? this._buildRequestKeyFromItemId(parseInt(item.ID.toString(), 10)) : ""),
          requesterName: item[employeeKey] || item.Title || "",
          assetId: "",
          assetTitle: item[selectAssetKey] || item.Title || "",
          quantity: parseInt(item[quantityKey]) || 1,
          status,
          assetStatus: (item[assetStatusKey] || "Pending").toString().toLowerCase().includes("approv") ? "Approved" : "Pending",
          managerResponse: item[managerCommentKey] || "",
          requestDate: item.Created ? item.Created.split('T')[0] : new Date().toISOString().split('T')[0],
          reason: item[reasonKey] || ""
        };
      });
    } catch (error: any) {
      console.error("Error fetching requests from SharePoint:", error);
      throw error;
    }
  }

  public static async updateRequestStatus(
    requestId: number,
    status: 'Approved' | 'Declined',
    approverName: string = 'Unknown',
    rejectionReason?: string
  ): Promise<void> {
    const sp = getSP();
    try {
      await this._ensureRequestWorkflowFields();

      if (Number.isNaN(requestId)) {
        throw new Error('Invalid request ID');
      }

      const list = sp.web.lists.getByTitle(InventoryService.REQUEST_LIST_NAME);
      const item = await list.items.getById(requestId).select("*")();
      const keys = Object.keys(item || {});
      const findKey = (searchStr: string): string | undefined =>
        keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').includes(searchStr));

      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString", "Choices")();
      const statusField = fields.find(field => {
        const internalNameRaw = (field.InternalName || '').toString();
        const internalName = internalNameRaw.toLowerCase();
        const title = ((field.Title || '') as string).toLowerCase();
        const normalizedInternal = internalName.replace(/_x0020_/g, '');
        const isModerationField = internalName.includes('moderation');
        const isBusinessStatusField = normalizedInternal === 'status' || title.trim() === 'status';
        return isBusinessStatusField && !isModerationField;
      });

      const statusKeyFromItem = keys.find(key => this._isBusinessStatusKey(key));

      const statusKey = statusKeyFromItem || statusField?.InternalName || InventoryService.REQUEST_STATUS_INTERNAL_NAME;
      if (!statusKey) {
        throw new Error('Could not find request status column. Please create a Choice column like RequestStatus/Status in RequestList.');
      }

      if (!this._isBusinessStatusKey(statusKey)) {
        throw new Error('Detected non-business status field. Please ensure RequestList has a dedicated request status column.');
      }
      const reasonKey = findKey("managercomment") || InventoryService.REQUEST_COMMENT_INTERNAL_NAME || findKey("rejectionreason") || findKey("comments") || findKey("reason");
      const rawChoices = statusField?.Choices;
      const choices: string[] = Array.isArray(rawChoices)
        ? rawChoices
        : (rawChoices && Array.isArray(rawChoices.results) ? rawChoices.results : []);

      const pickChoice = (preferred: string[], fallback: string): string => {
        if (!choices.length) {
          return fallback;
        }

        const lowerChoices = choices.map(choice => (choice || '').toString().toLowerCase());
        for (const preferredValue of preferred) {
          const preferredLower = preferredValue.toLowerCase();
          for (let i = 0; i < lowerChoices.length; i++) {
            if (lowerChoices[i].includes(preferredLower) || preferredLower.includes(lowerChoices[i])) {
              return choices[i];
            }
          }
        }

        return fallback;
      };

      const statusValue = status === 'Declined'
        ? pickChoice(['rejected', 'declined'], 'Rejected')
        : pickChoice(['approved'], 'Approved');
      const requestKey = this._extractRequestKey(item);

      const basePayload: any = {};
      basePayload[statusKey] = statusValue;
      if (reasonKey) {
        basePayload[reasonKey] = status === 'Declined'
          ? (rejectionReason || 'Rejected by manager')
          : `Approved by ${approverName}`;
      }
      await list.items.getById(requestId).update(basePayload);

      await this.addAuditLog({
        title: `${statusValue} Request ${requestKey || `#${requestId}`}`,
        action: 'Update',
        entityType: 'Request',
        entityId: requestKey || requestId.toString(),
        details: JSON.stringify({
          requestKey: requestKey || this._buildRequestKeyFromItemId(requestId),
          lifecycle: statusValue,
          changedBy: approverName,
          changedAt: new Date().toISOString(),
          rejectionReason: status === 'Declined' ? (rejectionReason || "") : "",
          assetAllocation: status === 'Approved'
            ? {
                assetTitle: item[findKey("selectasset") || "SelectAsset"] || item.Title || "",
                quantity: parseInt(item[findKey("quantity") || "Quantity"], 10) || 1
              }
            : undefined
        }),
        user: approverName
      });
    } catch (error: any) {
      console.error(`Failed to update RequestList item ${requestId} status`, error);
      throw new Error(`Unable to update request status. ${error.message || 'Verify RequestList status column and choices.'}`);
    }
  }

  public static async addAuditLog(log: Omit<IEventLog, 'id' | 'timestamp'>): Promise<void> {
    const sp = getSP();
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
    }
  }

  public static async updateAssetStatus(
    requestId: number,
    assetStatus: 'Approved' | 'Pending',
    approverName: string = 'Unknown'
  ): Promise<void> {
    const sp = getSP();
    try {
      await this._ensureRequestWorkflowFields();

      if (Number.isNaN(requestId)) {
        throw new Error('Invalid request ID');
      }

      const list = sp.web.lists.getByTitle(InventoryService.REQUEST_LIST_NAME);
      const item = await list.items.getById(requestId).select("*")();
      const keys = Object.keys(item || {});
      const findKey = (searchStr: string): string | undefined =>
        keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').includes(searchStr));
      const requestKey = this._extractRequestKey(item);
      const assetStatusKey = findKey("assetstatus") || InventoryService.ASSET_STATUS_INTERNAL_NAME;

      await list.items.getById(requestId).update({
        [assetStatusKey]: assetStatus
      });

      await this.addAuditLog({
        title: `Asset status ${assetStatus} for Request ${requestKey || `#${requestId}`}`,
        action: 'Update',
        entityType: 'Request',
        entityId: requestKey || requestId.toString(),
        details: JSON.stringify({
          requestKey: requestKey || this._buildRequestKeyFromItemId(requestId),
          lifecycle: "AssetStatusUpdated",
          assetStatus,
          changedBy: approverName,
          changedAt: new Date().toISOString()
        }),
        user: approverName
      });
    } catch (error: any) {
      console.error(`Failed to update asset status for RequestList item ${requestId}`, error);
      throw new Error(`Unable to update asset status. ${error.message || 'Verify AssetStatus column and permissions.'}`);
    }
  }

  public static async assignAssetsToEmployee(
    
    assetIds: string[], 
    employeeName: string, 
    employeeEmail: string, 
    adminName: string
  ): Promise<void> {
    const sp = getSP();
    const list = sp.web.lists.getByTitle(InventoryService.LIST_NAME);
    let assignedToId: number | null = null;
    
    // Try to resolve the user in SharePoint by email
    try {
      const user: any = await sp.web.ensureUser(employeeEmail);
      assignedToId = user.data ? user.data.Id : user.Id;
    } catch (e) {
      console.warn(`Could not resolve user ${employeeEmail} in SharePoint. Falling back to string assignment if column allows.`, e);
    }

    const updatePromises = assetIds.map(async (assetId) => {
      const payloadsToTry: any[] = [];
      const baseStatus = { Status: 'Assigned' };
      
      if (assignedToId !== null) {
        payloadsToTry.push({ ...baseStatus, AssignedToId: { results: [assignedToId] }, Note: `Assigned to: ${employeeName}` });
        payloadsToTry.push({ ...baseStatus, Assigned_x0020_ToId: { results: [assignedToId] }, Note: `Assigned to: ${employeeName}` });
        payloadsToTry.push({ ...baseStatus, AssignedToId: assignedToId, Note: `Assigned to: ${employeeName}` });
        payloadsToTry.push({ ...baseStatus, Assigned_x0020_ToId: assignedToId, Note: `Assigned to: ${employeeName}` });
        payloadsToTry.push({ ...baseStatus, AssignedToId: assignedToId });
        payloadsToTry.push({ ...baseStatus, Assigned_x0020_ToId: assignedToId });
      } else {
        payloadsToTry.push({ ...baseStatus, AssignedTo: employeeName, Note: `Assigned to: ${employeeName}` });
        payloadsToTry.push({ ...baseStatus, Assigned_x0020_To: employeeName, Note: `Assigned to: ${employeeName}` });
        payloadsToTry.push({ ...baseStatus, AssignedTo: employeeName });
        payloadsToTry.push({ ...baseStatus, Assigned_x0020_To: employeeName });
      }
      
      payloadsToTry.push({ ...baseStatus, Note: `Assigned to: ${employeeName}` });
      payloadsToTry.push({ ...baseStatus });

      let success = false;
      let lastErr: any;

      for (const payload of payloadsToTry) {
        try {
          await list.items.getById(parseInt(assetId)).update(payload);
          success = true;
          break; // Stop trying if one succeeds
        } catch (err) {
          lastErr = err;
        }
      }

      if (!success) {
        console.error(`All fallback updates failed for asset ${assetId}`, lastErr);
        throw new Error(lastErr.message || "Failed to update asset status");
      }

      await this.addAuditLog({
        title: `Asset directly assigned to ${employeeName}`,
        action: 'Update',
        entityType: 'Asset',
        entityId: assetId,
        details: JSON.stringify({
          lifecycle: "DirectAssignment",
          assignedTo: employeeName,
          changedBy: adminName,
          changedAt: new Date().toISOString()
        }),
        user: adminName
      });
    });

    await Promise.all(updatePromises);
  }

  public static async getAuditLogs(): Promise<IEventLog[]> {
    const sp = getSP();
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

  public static async getRequestHistoryById(
    requestLookupId: string
  ): Promise<{ request: IRequest; lifecycle: IEventLog[] }> {
    const sp = getSP();
    await this._ensureRequestWorkflowFields();

    const normalizedRequestKey = this._normalizeRequestKey(requestLookupId);
    if (!normalizedRequestKey) {
      throw new Error("Request ID is required.");
    }

    let requestItems: any[] = [];
    try {
      requestItems = await sp.web.lists
        .getByTitle(InventoryService.REQUEST_LIST_NAME)
        .items.select("*")
        .filter(`${InventoryService.REQUEST_KEY_INTERNAL_NAME} eq '${normalizedRequestKey.replace(/'/g, "''")}'`)();
    } catch (filterError) {
      // RequestKey column might not exist yet; fallback to derived ID format.
      console.warn("RequestKey filter failed. Falling back to item ID based lookup.", filterError);
    }

    if (!requestItems.length) {
      const derivedIdMatch = /^REQ-(\d{1,})$/.exec(normalizedRequestKey.replace(/^REQ-0*/, "REQ-"));
      const parsedId = derivedIdMatch ? parseInt(derivedIdMatch[1], 10) : NaN;
      if (!Number.isNaN(parsedId)) {
        try {
          const requestById = await sp.web.lists
            .getByTitle(InventoryService.REQUEST_LIST_NAME)
            .items.getById(parsedId)
            .select("*")();
          requestItems = requestById ? [requestById] : [];
        } catch (err) {
          console.warn(`Fallback ID lookup failed for ${normalizedRequestKey}.`, err);
        }
      }
    }

    if (!requestItems || requestItems.length === 0) {
      throw new Error(`No request found for ID ${normalizedRequestKey}`);
    }

    const requestItem = requestItems[0];
    const requests = await this.getRequests();
    const request = requests.find(r =>
      this._normalizeRequestKey(r.requestKey) === normalizedRequestKey ||
      r.id === requestItem.ID?.toString()
    );

    if (!request) {
      throw new Error(`Request exists but could not be mapped for ID ${normalizedRequestKey}`);
    }

    const requestIdAsString = requestItem.ID ? requestItem.ID.toString() : "";
    const allLogs = await this.getAuditLogs();
    const lifecycle = allLogs
      .filter(log =>
        log.entityType === "Request" && (
          this._normalizeRequestKey(log.entityId) === normalizedRequestKey ||
          log.entityId === requestIdAsString ||
          ((log.details || "").toUpperCase().indexOf(`"REQUESTKEY":"${normalizedRequestKey}"`) >= 0)
        )
      )
      .sort((a, b) => new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime());

    return {
      request,
      lifecycle
    };
  }
}
