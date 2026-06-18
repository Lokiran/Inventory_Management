import { IInventoryItem } from "../models/IInventoryItem";
import { IRequest } from "../models/IRequest";
import { IEventLog } from "../models/IEventLog";
import { IReturnRequest } from "../models/IReturnRequest";
import { getSP } from "../pnpjsConfig";
import { EMPLOYEES } from "../data/mockData";

export class InventoryService {
  private static readonly LIST_NAME = "InventoryList";
  private static readonly EVENT_LOG_LIST = "EventLogList";
  private static readonly REQUEST_LIST_NAME = "RequestList";
  private static readonly REQUEST_STATUS_INTERNAL_NAME = "RequestStatus";
  private static readonly REQUEST_COMMENT_INTERNAL_NAME = "ManagerComment";
  private static readonly REQUEST_KEY_INTERNAL_NAME = "RequestKey";
  private static readonly ASSET_STATUS_INTERNAL_NAME = "AssetStatus";
  private static readonly MAPPING_LIST_NAME = "Mapping List";

  private static _requestWorkflowFieldsEnsured = false;
  private static _resolvedListName: string | null = null;
  private static _resolvedRequestListName: string | null = null;
  private static _resolvedMappingListName: string | null = null;
  private static _mappingListFieldsEnsured = false;

  public static async getInventoryList(): Promise<any> {
    const sp = getSP();
    if (InventoryService._resolvedListName) {
      return sp.web.lists.getByTitle(InventoryService._resolvedListName);
    }
    try {
      const list = sp.web.lists.getByTitle(InventoryService.LIST_NAME);
      InventoryService._resolvedListName = InventoryService.LIST_NAME; // Assign before await to satisfy require-atomic-updates
      await list.select("Title")(); // Verify list exists
      return list;
    } catch (e) {
      InventoryService._resolvedListName = null;
      try {
        const fallbackName = "Inventory List";
        const list = sp.web.lists.getByTitle(fallbackName);
        InventoryService._resolvedListName = fallbackName; // Assign before await to satisfy require-atomic-updates
        await list.select("Title")(); // Verify fallback exists
        console.log("Resolved list name dynamically to fallback: " + fallbackName);
        return list;
      } catch (e2) {
        InventoryService._resolvedListName = null;
        try {
          const allLists = await sp.web.lists.select("Title")();
          const listNames = allLists.map(l => '"' + l.Title + '"').join(', ');
          throw new Error("List '" + InventoryService.LIST_NAME + "' or 'Inventory List' does not exist on this SharePoint site. Available lists on this site are: [ " + listNames + " ]. Please ensure your list title matches exactly.");
        } catch (listsError) {
          throw new Error("List '" + InventoryService.LIST_NAME + "' or 'Inventory List' does not exist on this SharePoint site.");
        }
      }
    }
  }

  public static async getRequestList(): Promise<any> {
    const sp = getSP();
    if (InventoryService._resolvedRequestListName) {
      return sp.web.lists.getByTitle(InventoryService._resolvedRequestListName);
    }
    try {
      const list = sp.web.lists.getByTitle(InventoryService.REQUEST_LIST_NAME);
      InventoryService._resolvedRequestListName = InventoryService.REQUEST_LIST_NAME; // Assign before await to satisfy require-atomic-updates
      await list.select("Title")(); // Verify list exists
      return list;
    } catch (e) {
      InventoryService._resolvedRequestListName = null;
      try {
        const fallbackName = "Request List";
        const list = sp.web.lists.getByTitle(fallbackName);
        InventoryService._resolvedRequestListName = fallbackName; // Assign before await to satisfy require-atomic-updates
        await list.select("Title")(); // Verify fallback exists
        console.log("Resolved requests list name dynamically to fallback: " + fallbackName);
        return list;
      } catch (e2) {
        InventoryService._resolvedRequestListName = null;
        try {
          const allLists = await sp.web.lists.select("Title")();
          const listNames = allLists.map(l => '"' + l.Title + '"').join(', ');
          throw new Error("List '" + InventoryService.REQUEST_LIST_NAME + "' or 'Request List' does not exist on this SharePoint site. Available lists are: [ " + listNames + " ].");
        } catch (listsError) {
          throw new Error("List '" + InventoryService.REQUEST_LIST_NAME + "' or 'Request List' does not exist.");
        }
      }
    }
  }

  public static async getMappingList(): Promise<any> {
    const sp = getSP();
    if (InventoryService._resolvedMappingListName) {
      return sp.web.lists.getByTitle(InventoryService._resolvedMappingListName);
    }
    try {
      const list = sp.web.lists.getByTitle(InventoryService.MAPPING_LIST_NAME);
      InventoryService._resolvedMappingListName = InventoryService.MAPPING_LIST_NAME;
      await list.select("Title")(); // Verify list exists
      return list;
    } catch (e) {
      InventoryService._resolvedMappingListName = null;
      try {
        const fallbackName = "MappingList";
        const list = sp.web.lists.getByTitle(fallbackName);
        InventoryService._resolvedMappingListName = fallbackName;
        await list.select("Title")(); // Verify fallback exists
        console.log("Resolved mapping list name dynamically to fallback: " + fallbackName);
        return list;
      } catch (e2) {
        InventoryService._resolvedMappingListName = null;
        // Attempt to auto-create "Mapping List" dynamically
        try {
          console.log("Attempting to auto-create 'Mapping List' list...");
          await sp.web.lists.add(InventoryService.MAPPING_LIST_NAME, "List for tracking asset assignments", 100);
          InventoryService._resolvedMappingListName = InventoryService.MAPPING_LIST_NAME;
          console.log("Successfully created 'Mapping List' in SharePoint.");
          return sp.web.lists.getByTitle(InventoryService.MAPPING_LIST_NAME);
        } catch (createError) {
          InventoryService._resolvedMappingListName = null;
          try {
            const allLists = await sp.web.lists.select("Title")();
            const listNames = allLists.map(l => '"' + l.Title + '"').join(', ');
            throw new Error("List '" + InventoryService.MAPPING_LIST_NAME + "' does not exist and could not be auto-created on this SharePoint site. Available lists are: [ " + listNames + " ].");
          } catch (listsError) {
            throw new Error("List '" + InventoryService.MAPPING_LIST_NAME + "' does not exist and could not be auto-created.");
          }
        }
      }
    }
  }

  private static async _ensureMappingListFields(): Promise<void> {
    if (this._mappingListFieldsEnsured) {
      return;
    }

    try {
      const list = await InventoryService.getMappingList();
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString")();

      const hasField = (name: string) => fields.some(field => {
        const internalName = (field.InternalName || '').toString().toLowerCase();
        const title = (field.Title || '').toString().toLowerCase();
        const search = name.toLowerCase();
        return internalName === search || title === search;
      });

      // 1. Employe (Text)
      if (!hasField("Employe")) {
        try {
          await list.fields.addText("Employe");
        } catch (err) {
          console.warn("Could not auto-create Employe field. Continuing.", err);
        }
      }

      // 2. EmployeeID / Employee ID (Text)
      const hasEmpId = fields.some(field => {
        const val = (field.InternalName || '').toString().toLowerCase();
        const t = (field.Title || '').toString().toLowerCase();
        return val === 'employeeid' || val === 'employee_x0020_id' || val === 'employee id' || t === 'employee id' || t === 'employeeid' || val === 'employeid' || t === 'employeid' || t === 'employe id';
      });
      if (!hasEmpId) {
        try {
          await list.fields.addText("EmployeeID", { Title: "Employe ID" });
        } catch (err) {
          try {
            await list.fields.addText("EmployeID", { Title: "Employe ID" });
          } catch (err2) {
            try {
              await list.fields.addText("EmployeeID");
            } catch (err3) {
              console.warn("Could not auto-create EmployeeID field. Continuing.", err3);
            }
          }
        }
      }

      // 3. AssetName / Asset Name (Text)
      const hasAssetName = fields.some(field => {
        const val = (field.InternalName || '').toString().toLowerCase();
        const t = (field.Title || '').toString().toLowerCase();
        return val === 'assetname' || val === 'asset_x0020_name' || val === 'asset name' || t === 'asset name' || t === 'assetname';
      });
      if (!hasAssetName) {
        try {
          await list.fields.addText("AssetName", { Title: "Asset Name" });
        } catch (err) {
          try {
            await list.fields.addText("AssetName");
          } catch (err2) {
            console.warn("Could not auto-create AssetName field. Continuing.", err2);
          }
        }
      }

      // 4. SerialNumber / Serial Number (Text)
      const hasSerialNumber = fields.some(field => {
        const val = (field.InternalName || '').toString().toLowerCase();
        const t = (field.Title || '').toString().toLowerCase();
        return val === 'serialnumber' || val === 'serial_x0020_number' || val === 'serial number' || t === 'serial number' || t === 'serialnumber';
      });
      if (!hasSerialNumber) {
        try {
          await list.fields.addText("SerialNumber", { Title: "Serial Number" });
        } catch (err) {
          try {
            await list.fields.addText("SerialNumber");
          } catch (err2) {
            console.warn("Could not auto-create SerialNumber field. Continuing.", err2);
          }
        }
      }

      // 5. Priority (Choice)
      if (!hasField("Priority")) {
        try {
          await list.fields.addChoice("Priority", {
            Choices: ["High", "Medium", "Low"],
            FillInChoice: false
          });
        } catch (err) {
          console.warn("Could not auto-create Priority field. Continuing.", err);
        }
      }

      // 6. RequestedDate / Requested Date (Text)
      const hasRequestedDate = fields.some(field => {
        const val = (field.InternalName || '').toString().toLowerCase();
        const t = (field.Title || '').toString().toLowerCase();
        return val === 'requesteddate' || val === 'requested_x0020_date' || val === 'requested date' || t === 'requested date' || t === 'requesteddate';
      });
      if (!hasRequestedDate) {
        try {
          await list.fields.addText("RequestedDate", { Title: "Requested Date" });
        } catch (err) {
          try {
            await list.fields.addText("RequestedDate");
          } catch (err2) {
            console.warn("Could not auto-create RequestedDate field. Continuing.", err2);
          }
        }
      }

      // 7. ReasonforRequest / Reason for Request (Multiline Text)
      const hasReasonforRequest = fields.some(field => {
        const val = (field.InternalName || '').toString().toLowerCase();
        const t = (field.Title || '').toString().toLowerCase();
        return val === 'reasonforrequest' || val === 'reason_x0020_for_x0020_request' || val === 'reason for request' || t === 'reason for request' || t === 'reasonforrequest';
      });
      if (!hasReasonforRequest) {
        try {
          await list.fields.addMultilineText("ReasonforRequest", { Title: "Reason for Request" });
        } catch (err) {
          try {
            await list.fields.addMultilineText("ReasonforRequest");
          } catch (err2) {
            console.warn("Could not auto-create ReasonforRequest field. Continuing.", err2);
          }
        }
      }

      // 8. AssignedDate / Assigned Date (Text)
      const hasAssignedDate = fields.some(field => {
        const val = (field.InternalName || '').toString().toLowerCase();
        const t = (field.Title || '').toString().toLowerCase();
        return val === 'assigneddate' || val === 'assigned_x0020_date' || val === 'assigned date' || t === 'assigned date' || t === 'assigneddate';
      });
      if (!hasAssignedDate) {
        try {
          await list.fields.addText("AssignedDate", { Title: "Assigned Date" });
        } catch (err) {
          try {
            await list.fields.addText("AssignedDate");
          } catch (err2) {
            console.warn("Could not auto-create AssignedDate field. Continuing.", err2);
          }
        }
      }

    } catch (error) {
      console.warn("Could not ensure Mapping List fields. Continuing.", error);
    } finally {
      this._mappingListFieldsEnsured = true;
    }
  }

  private static _normalizeRequestKey(input: string): string {
    return (input || "").trim().toUpperCase();
  }

  private static _buildRequestKeyFromItemId(itemId: number): string {
    const raw = itemId.toString();
    const padded = ("000000" + raw).slice(-6);
    return `REQ-${padded}`;
  }

  private static _resolveRequestKeyInternalName(fields: any[]): string {
    const candidates = ["requestid", "requestkey", "request_x0020_id", "request_x0020_key", "request id"];
    for (const cand of candidates) {
      const field = fields.find((f: any) => {
        const internal = (f.InternalName || "").toLowerCase();
        const title = (f.Title || "").toLowerCase().replace(/[^a-z0-9]/g, "");
        const candNorm = cand.replace(/[^a-z0-9]/g, "");
        return internal === cand || internal.replace(/_x0020_/g, "") === candNorm || title === candNorm;
      });
      if (field) {
        return field.InternalName;
      }
    }
    return InventoryService.REQUEST_KEY_INTERNAL_NAME;
  }

  private static async _updateMissingRequestKeys(list: any, resolvedKeyName: string, items: any[]): Promise<void> {
    for (const item of items) {
      try {
        const itemId = parseInt(item.ID.toString(), 10);
        if (!Number.isNaN(itemId)) {
          const requestKey = this._buildRequestKeyFromItemId(itemId);
          await list.items.getById(itemId).update({
            [resolvedKeyName]: requestKey
          });
          console.log(`Successfully populated Request ID in SharePoint for item ${itemId}: ${requestKey}`);
        }
      } catch (err) {
        console.warn(`Failed to update missing Request ID for item ${item.ID}:`, err);
      }
    }
  }

  private static _extractRequestKey(item: any): string {
    if (!item) {
      return "";
    }

    const candidates = ["requestkey", "requestid", "request_x0020_id", "request_x0020_key"];
    for (const key of Object.keys(item)) {
      const normalizedKey = key.toLowerCase().replace(/_x0020_/g, "");
      if (candidates.indexOf(normalizedKey) >= 0 && item[key]) {
        return this._normalizeRequestKey(item[key].toString());
      }
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
    if (this._requestWorkflowFieldsEnsured) {
      return;
    }

    try {
      const list = await InventoryService.getRequestList();
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
        const name = (field.InternalName || '').toString().toLowerCase();
        const title = (field.Title || '').toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        return name === 'requestkey' || name === 'requestid' || name === 'request_x0020_id' || title === 'requestid' || title === 'requestkey';
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

      const hasEmployeeIdField = fields.some(field => {
        const internalName = (field.InternalName || '').toString().toLowerCase();
        return internalName === 'employeeid' || internalName === 'employee_x0020_id';
      });
      if (!hasEmployeeIdField) {
        try {
          await list.fields.addText('EmployeeID');
        } catch (err) {
          console.warn("Could not auto-create EmployeeID field. Continuing.", err);
        }
      }

      const hasPriorityField = fields.some(field => {
        const internalName = (field.InternalName || '').toString().toLowerCase();
        return internalName === 'priority';
      });
      if (!hasPriorityField) {
        try {
          await list.fields.addChoice('Priority', {
            Choices: ["High", "Medium", "Low"],
            FillInChoice: false
          });
        } catch (err) {
          console.warn("Could not auto-create Priority field. Continuing.", err);
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
    try {
      const list = await InventoryService.getInventoryList();
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString")();
      const items = await this._fetchItemsWithExpandedUsers(list);

      const findFieldInternalName = (searchStr: string, fallback: string): string => {
        let field = fields.find((f: any) => f.InternalName.toLowerCase() === searchStr.toLowerCase());
        if (field) return field.InternalName;
        field = fields.find((f: any) => f.InternalName.toLowerCase().replace(/_x0020_/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        if (field) return field.InternalName;
        field = fields.find((f: any) => f.Title.toLowerCase().replace(/[^a-z0-9]/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        return field ? field.InternalName : fallback;
      };

      const assignedToKey = findFieldInternalName("assignedto", "AssignedTo");
      const assetNameKey = findFieldInternalName("assetname", "AssetName");
      const assetTypeKey = findFieldInternalName("assettype", "AssetType");
      const serialNumberKey = findFieldInternalName("serialnumber", "SerialNumber");
      const purchaseDateKey = findFieldInternalName("purchasedate", "PurchaseDate");
      const vendorKey = findFieldInternalName("vendor", "Vendor");
      const conditionKey = findFieldInternalName("condition", "Condition");
      const statusKey = findFieldInternalName("status", "Status");
      const warrantyExpiryKey = findFieldInternalName("warrantyexpiry", "WarrantyExpiry");
      const specificationsKey = findFieldInternalName("specifications", "Specifications");
      const noteKey = findFieldInternalName("note", "Note");

      return items.map((item: any) => {
        const assignedToVal = (() => {
          const assignedField = item[assignedToKey] || item.AssignedTo || item.Assigned_x0020_To;
          if (assignedField) {
            if (typeof assignedField === 'string') return assignedField;
            if (Array.isArray(assignedField)) return assignedField.map((a: any) => a.Title).join(', ');
            if (assignedField.Title) return assignedField.Title;
          }

          // Fallback: extract assignee name from Note or Status if the primary field is empty
          const noteText = item[noteKey] || item.Note || item.Notes || "";
          const statusText = item[statusKey] || item.Status || item.AssetStatus || "";

          const extractFromText = (text: string): string => {
            const match = /assigned to:\s*([^;\n\r]+)/i.exec(text);
            return match ? match[1].trim() : "";
          };

          const fromNote = extractFromText(noteText);
          if (fromNote) return fromNote;

          const fromStatus = extractFromText(statusText);
          if (fromStatus) return fromStatus;

          return "";
        })();

        const rawStatus = item[statusKey] || item.Status || item.AssetStatus || "";
        let finalStatus = rawStatus;
        const statusLower = (rawStatus || "").toLowerCase();
        if (statusLower === "assigned" || statusLower.startsWith("assigned")) {
          if (!assignedToVal || assignedToVal.trim() === "") {
            finalStatus = "In Stock";
          }
        }

        return {
          id: item.ID.toString(),
          title: item.Title || "",
          assetName: item[assetNameKey] || item.AssetName || item.Asset_x0020_Name || item.Asset || "",
          assetType: item[assetTypeKey] || item.AssetType || item.Asset_x0020_Type || item.Type || "",
          serialNumber: item[serialNumberKey] || item.SerialNumber || item.Serial_x0020_Number || "",
          purchaseDate: item[purchaseDateKey] || item.PurchaseDate || item.Purchase_x0020_Date || "",
          vendor: item[vendorKey] || item.Vendor || item.VendorName || "",
          condition: item[conditionKey] || item.Condition || item.AssetCondition || "",
          status: finalStatus,
          assignedTo: assignedToVal,
          assignedDate: item.Modified || "",
          warrantyExpiry: item[warrantyExpiryKey] || item.WarrantyExpiry || item.Warranty_x0020_Expiry || "",
          specifications: item[specificationsKey] || item.Specifications || item.SpecificationsText || item.Note || item.Notes || "",
          note: item[noteKey] || item.Note || item.Notes || ""
        };
      });
    } catch (error: any) {
      console.error("Error fetching items from SharePoint:", error);
      throw error;
    }
  }

  public static async addItem(item: Omit<IInventoryItem, 'id'>, userDisplayName: string = "Unknown"): Promise<void> {
    const list = await InventoryService.getInventoryList();
    const payloads = [
      // 1. Standard modern field names with Specifications column (Priority)
      {
        Title: item.title,
        AssetName: item.assetName,
        AssetType: item.assetType,
        SerialNumber: item.serialNumber,
        PurchaseDate: item.purchaseDate,
        Vendor: item.vendor || "",
        Condition: item.condition || "",
        WarrantyExpiry: item.warrantyExpiry || "",
        Status: item.status,
        Specifications: item.specifications || ""
      },
      // 2. Space field names with Specifications column
      {
        Title: item.title,
        Asset_x0020_Name: item.assetName,
        Asset_x0020_Type: item.assetType,
        Serial_x0020_Number: item.serialNumber,
        Purchase_x0020_Date: item.purchaseDate,
        Vendor: item.vendor || "",
        Condition: item.condition || "",
        WarrantyExpiry: item.warrantyExpiry || "",
        Status: item.status,
        Specifications: item.specifications || ""
      },
      // 3. Alternate field names with Specifications column
      {
        Title: item.title,
        Asset: item.assetName,
        Type: item.assetType,
        Serial_x0020_Number: item.serialNumber,
        Purchase_x0020_Date: item.purchaseDate,
        Vendor: item.vendor || "",
        Condition: item.condition || "",
        WarrantyExpiry: item.warrantyExpiry || "",
        AssetStatus: item.status,
        Specifications: item.specifications || ""
      },
      // 4. Standard modern field names fallback (Note)
      {
        Title: item.title,
        AssetName: item.assetName,
        AssetType: item.assetType,
        SerialNumber: item.serialNumber,
        PurchaseDate: item.purchaseDate,
        Vendor: item.vendor || "",
        Condition: item.condition || "",
        WarrantyExpiry: item.warrantyExpiry || "",
        Status: item.status,
        Note: item.specifications || ""
      },
      // 5. Space field names fallback (Note)
      {
        Title: item.title,
        Asset_x0020_Name: item.assetName,
        Asset_x0020_Type: item.assetType,
        Serial_x0020_Number: item.serialNumber,
        Purchase_x0020_Date: item.purchaseDate,
        Vendor: item.vendor || "",
        Condition: item.condition || "",
        WarrantyExpiry: item.warrantyExpiry || "",
        Status: item.status,
        Note: item.specifications || ""
      },
      // 6. Alternate field names fallback (Notes)
      {
        Title: item.title,
        Asset: item.assetName,
        Type: item.assetType,
        Serial_x0020_Number: item.serialNumber,
        Purchase_x0020_Date: item.purchaseDate,
        Vendor: item.vendor || "",
        Condition: item.condition || "",
        WarrantyExpiry: item.warrantyExpiry || "",
        AssetStatus: item.status,
        Notes: item.specifications || ""
      }
    ];

    let addedItem: any;
    let success = false;
    let lastError: any;
    for (const payload of payloads) {
      try {
        addedItem = await list.items.add(payload);
        success = true;
        break; // Success, stop looping immediately!
      } catch (error) {
        lastError = error;
      }
    }

    if (!success) {
      console.error("Error adding item to SharePoint:", lastError);
      throw new Error(`SharePoint rejected the save. The columns you created in InventoryList do not match the expected format. Error: ${lastError.message || JSON.stringify(lastError)}`);
    }

    // Safely perform post-save actions (audit logging) outside the creation loop
    try {
      const entityId = (addedItem && addedItem.data && addedItem.data.Id)
        ? addedItem.data.Id.toString()
        : (addedItem && addedItem.Id)
          ? addedItem.Id.toString()
          : 'Unknown';

      await this.addAuditLog({
        title: `Created Asset: ${item.assetName || item.title} (Inactivated)`,
        action: 'Inactivated',
        entityType: 'Asset',
        entityId,
        details: JSON.stringify({
          lifecycle: "Registration",
          status: "In Stock",
          changedBy: userDisplayName,
          changedAt: new Date().toISOString(),
          item
        }),
        user: userDisplayName
      });
    } catch (auditError) {
      console.warn("Failed to write audit log for newly created asset:", auditError);
    }
  }

  public static async addRequest(request: Omit<IRequest, 'id' | 'requestKey' | 'status'> & { status?: string }, userDisplayName: string = "Unknown"): Promise<void> {
    const list = await InventoryService.getRequestList();
    await this._ensureRequestWorkflowFields();

    const initialStatus = request.status || "Pending";
    const sp = getSP();
    let requesterId: number | null = null;
    try {
      const user: any = await sp.web.ensureUser(request.requesterName);
      requesterId = user.data ? user.data.Id : user.Id;
    } catch (e) {
      console.warn("Could not resolve requester in SharePoint", e);
    }

    let dynamicPayload: any = null;
    try {
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString", "Required")();

      const findField = (searchStr: string) => {
        let field = fields.find((f: any) => f.InternalName.toLowerCase() === searchStr.toLowerCase());
        if (field) return field;
        field = fields.find((f: any) => f.InternalName.toLowerCase().replace(/_x0020_/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        if (field) return field;
        field = fields.find((f: any) => f.Title.toLowerCase().replace(/[^a-z0-9]/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        return field;
      };

      const requesterField = findField("employee") || findField("requester");
      const assetField = findField("assettype") || findField("asset type") || findField("selectasset") || findField("asset") || findField("type");
      const quantityField = findField("quantity");
      const reasonField = findField("reason");
      const statusField = fields.find((f: any) => {
        const name = (f.InternalName || '').toLowerCase().replace(/_x0020_/g, '');
        const title = (f.Title || '').toLowerCase().replace(/[^a-z0-9]/g, '');
        return name === "requeststatus" || name === "status" || title === "requeststatus" || title === "status";
      });
      const employeeIdField = findField("employeeid") || findField("employee id") || findField("employee_x0020_id");
      const priorityField = findField("priority");
      const requestDateField = findField("requestdate") || findField("request date") || findField("requesteddate") || findField("requested date");

      dynamicPayload = {
        Title: `Request for ${request.assetTitle}`
      };

      if (requesterField) {
        const isPerson = requesterField.TypeAsString === "User" || requesterField.TypeAsString === "UserMulti";
        if (isPerson && requesterId !== null) {
          dynamicPayload[`${requesterField.InternalName}Id`] = requesterId;
        } else {
          dynamicPayload[requesterField.InternalName] = request.requesterName;
        }
      }

      if (assetField) {
        const isLookup = assetField.TypeAsString === "Lookup";
        if (isLookup) {
          dynamicPayload[`${assetField.InternalName}Id`] = parseInt(request.assetId, 10) || 1;
        } else {
          dynamicPayload[assetField.InternalName] = request.assetTitle;
        }
      }

      if (quantityField) {
        dynamicPayload[quantityField.InternalName] = request.quantity;
      }

      if (reasonField) {
        dynamicPayload[reasonField.InternalName] = request.reason || "";
      }

      if (statusField) {
        dynamicPayload[statusField.InternalName] = initialStatus;
      } else {
        dynamicPayload[InventoryService.REQUEST_STATUS_INTERNAL_NAME] = initialStatus;
      }

      if (employeeIdField) {
        dynamicPayload[employeeIdField.InternalName] = (request as any).employeeId || "";
      }
      if (priorityField) {
        dynamicPayload[priorityField.InternalName] = (request as any).priority || "Medium";
      }
      if (requestDateField) {
        dynamicPayload[requestDateField.InternalName] = request.requestDate || new Date().toISOString().split('T')[0];
      }
    } catch (e) {
      console.warn("Failed to generate dynamic payload from schema, will use hardcoded candidates", e);
    }

    const payloads = [
      ...(dynamicPayload ? [dynamicPayload] : []),
      // 1. User's specific columns (Employee/Requester as Person/Lookup field)
      ...(requesterId !== null ? [
        {
          Title: `Request for ${request.assetTitle}`,
          EmployeeId: requesterId,
          EmployeeID: (request as any).employeeId || "",
          Priority: (request as any).priority || "Medium",
          Asset_x0020_type: request.assetTitle,
          Quantity: request.quantity,
          Reason_x0020_for_x0020_Request: request.reason || "",
          RequestStatus: initialStatus,
          RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
          Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
        },
        {
          Title: `Request for ${request.assetTitle}`,
          EmployeeId: requesterId,
          EmployeeID: (request as any).employeeId || "",
          Priority: (request as any).priority || "Medium",
          Assettype: request.assetTitle,
          Quantity: request.quantity,
          ReasonforRequest: request.reason || "",
          RequestStatus: initialStatus,
          RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
          Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
        },
        {
          Title: `Request for ${request.assetTitle}`,
          RequesterId: requesterId,
          EmployeeID: (request as any).employeeId || "",
          Priority: (request as any).priority || "Medium",
          Asset_x0020_type: request.assetTitle,
          Quantity: request.quantity,
          Reason_x0020_for_x0020_Request: request.reason || "",
          RequestStatus: initialStatus,
          RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
          Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
        },
        {
          Title: `Request for ${request.assetTitle}`,
          RequesterId: requesterId,
          EmployeeID: (request as any).employeeId || "",
          Priority: (request as any).priority || "Medium",
          Assettype: request.assetTitle,
          Quantity: request.quantity,
          ReasonforRequest: request.reason || "",
          RequestStatus: initialStatus,
          RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
          Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
        }
      ] : []),
      // 2. User's specific columns (Employee/Requester as plain Text field)
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Asset_x0020_type: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
        RequestStatus: initialStatus,
        RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
        Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Assettype: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: initialStatus,
        RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
        Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Requester: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Asset_x0020_type: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
        RequestStatus: initialStatus,
        RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
        Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Requester: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Assettype: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: initialStatus,
        RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
        Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
      },
      // 3. Fallbacks
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        SelectAsset: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: initialStatus,
        RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
        Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Select_x0020_Asset: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
        RequestStatus: initialStatus,
        RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
        Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: initialStatus,
        RequestDate: request.requestDate || new Date().toISOString().split('T')[0],
        Request_x0020_Date: request.requestDate || new Date().toISOString().split('T')[0]
      }
    ];

    let addedRequest: any;
    let success = false;
    let lastError: any;
    for (const payload of payloads) {
      try {
        addedRequest = await list.items.add(payload);
        success = true;
        break; // Success, stop looping immediately!
      } catch (err) {
        lastError = err;
      }
    }

    if (!success) {
      console.error("Error adding request to SharePoint after trying all column combinations.", lastError);
      throw new Error(`SharePoint rejected the save. The columns you created in RequestList do not match the expected format. Please check the Developer Console (F12) for the exact column name mismatch.`);
    }

    // Safely perform post-save actions (key generation, updating, logging) outside the creation loop
    try {
      const requestItemId = (addedRequest && addedRequest.data && addedRequest.data.Id)
        ? parseInt(addedRequest.data.Id.toString(), 10)
        : (addedRequest && addedRequest.Id)
          ? parseInt(addedRequest.Id.toString(), 10)
          : NaN;

      const requestKey = Number.isNaN(requestItemId)
        ? `REQ-${Date.now().toString(36).toUpperCase()}`
        : this._buildRequestKeyFromItemId(requestItemId);

      if (!Number.isNaN(requestItemId)) {
        try {
          const requestListInstance = await InventoryService.getRequestList();
          const fields: any[] = await requestListInstance.fields.select("InternalName", "Title")();
          const resolvedKeyName = InventoryService._resolveRequestKeyInternalName(fields);
          await requestListInstance.items.getById(requestItemId)
            .update({
              [resolvedKeyName]: requestKey,
              [InventoryService.ASSET_STATUS_INTERNAL_NAME]: "Pending"
            });
        } catch (err) {
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
    } catch (postError) {
      console.warn("Failed in post-request creation steps:", postError);
    }
  }

  public static async deleteItem(id: number, itemTitle: string = "Unknown", userDisplayName: string = "Unknown"): Promise<void> {
    try {
      const list = await InventoryService.getInventoryList();
      await list.items.getById(id).delete();

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
    try {
      await this._ensureRequestWorkflowFields();
      const list = await InventoryService.getRequestList();
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString")();
      const items = await this._fetchItemsWithExpandedUsers(list);

      const findFieldInternalName = (searchStr: string, fallback: string): string => {
        let field = fields.find((f: any) => f.InternalName.toLowerCase() === searchStr.toLowerCase());
        if (field) return field.InternalName;
        field = fields.find((f: any) => f.InternalName.toLowerCase().replace(/_x0020_/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        if (field) return field.InternalName;
        field = fields.find((f: any) => f.Title.toLowerCase().replace(/[^a-z0-9]/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        return field ? field.InternalName : fallback;
      };

      const employeeKey = findFieldInternalName("employee", "Employee");
      const requesterKey = findFieldInternalName("requester", "Requester");
      const selectAssetKey = findFieldInternalName("assettype", "SelectAsset");
      const quantityKey = findFieldInternalName("quantity", "Quantity");
      const reasonKey = findFieldInternalName("reason", "ReasonforRequest");
      const managerCommentKey = findFieldInternalName("managercomment", "ManagerComment");
      const assetStatusKey = findFieldInternalName("assetstatus", "AssetStatus");
      const statusKey = findFieldInternalName("requeststatus", "RequestStatus");

      const employeeIdKey = findFieldInternalName("employeeid", "EmployeeID");
      const priorityKey = findFieldInternalName("priority", "Priority");
      const requestDateKey = findFieldInternalName("requestdate", "RequestDate");

      const resolvedKeyName = this._resolveRequestKeyInternalName(fields);

      const mapped = items.map((item: any) => {
        const rawStatus = item[statusKey] || item.Status || 'Pending';
        const normalizedStatus = (rawStatus || '').toString().toLowerCase();
        const status: 'Pending' | 'Approved' | 'Declined' =
          (normalizedStatus.includes('approv')) ? 'Approved' :
            (normalizedStatus.includes('declin') || normalizedStatus.includes('reject')) ? 'Declined' :
              'Pending';
        const requestKey = item[resolvedKeyName] || this._extractRequestKey(item);

        return {
          id: item.ID ? item.ID.toString() : Math.random().toString(36).substr(2, 9),
          requestKey: requestKey || (item.ID ? this._buildRequestKeyFromItemId(parseInt(item.ID.toString(), 10)) : ""),
          requesterName: (() => {
            const rawEmp = item[employeeKey] || item[requesterKey] || item.Employee || item.Author;
            if (!rawEmp) return item.Title || "";
            if (typeof rawEmp === 'string') return rawEmp;
            if (Array.isArray(rawEmp)) return rawEmp.map((a: any) => a.Title || a.Name || "").join(', ');
            if (typeof rawEmp === 'object') return rawEmp.Title || rawEmp.Name || JSON.stringify(rawEmp);
            return rawEmp.toString();
          })(),
          employeeId: item[employeeIdKey] || "",
          assetId: "",
          assetTitle: item[selectAssetKey] || item.Title || "",
          assetName: "",
          priority: item[priorityKey] || "Medium",
          quantity: parseInt(item[quantityKey]) || 1,
          status,
          assetStatus: ((item[assetStatusKey] || "Pending").toString().toLowerCase().includes("approv") ? "Approved" : "Pending") as 'Pending' | 'Approved',
          managerResponse: item[managerCommentKey] || "",
          requestDate: item[requestDateKey] ? item[requestDateKey].split('T')[0] : (item.Created ? item.Created.split('T')[0] : new Date().toISOString().split('T')[0]),
          reason: item[reasonKey] || ""
        };
      });

      const itemsToUpdate = items.filter((item: any) => !item[resolvedKeyName] && item.ID);
      if (itemsToUpdate.length > 0) {
        this._updateMissingRequestKeys(list, resolvedKeyName, itemsToUpdate).catch(err => {
          console.warn("Background update of missing RequestKeys failed:", err);
        });
      }

      return mapped;
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
    try {
      await this._ensureRequestWorkflowFields();

      if (Number.isNaN(requestId)) {
        throw new Error('Invalid request ID');
      }

      const list = await InventoryService.getRequestList();
      const item = await list.items.getById(requestId).select("*")();
      const keys = Object.keys(item || {});
      const findKey = (searchStr: string): string | undefined => {
        const nonIdMatch = keys.find(k => {
          const kl = k.toLowerCase().replace(/_x0020_/g, '');
          return kl.indexOf(searchStr) >= 0 && !kl.endsWith("id");
        });
        if (nonIdMatch) return nonIdMatch;
        return keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').indexOf(searchStr) >= 0);
      };

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
              assetTitle: item[findKey("assettype") || findKey("selectasset") || findKey("type") || "SelectAsset"] || item.Title || "",
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
    try {
      await this._ensureRequestWorkflowFields();

      if (Number.isNaN(requestId)) {
        throw new Error('Invalid request ID');
      }

      const list = await InventoryService.getRequestList();
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

  private static _resolveMappingPayload(
    mappingFields: any[],
    employeeName: string,
    employeeId: string,
    assetName: string,
    serialNumber: string,
    priority: string,
    requestedDate: string,
    reason: string,
    assignedDate: string,
    assignedToId: number | null,
    isEmployePerson: boolean
  ): any {
    const findField = (searchStr: string, fallback: string): string => {
      let field = mappingFields.find((f: any) => f.Title.toLowerCase() === searchStr.toLowerCase());
      if (field) return field.InternalName;

      field = mappingFields.find((f: any) => f.InternalName.toLowerCase() === searchStr.toLowerCase());
      if (field) return field.InternalName;

      const normalizedSearch = searchStr.toLowerCase().replace(/[^a-z0-9]/g, '');
      field = mappingFields.find((f: any) => f.Title.toLowerCase().replace(/[^a-z0-9]/g, '') === normalizedSearch);
      if (field) return field.InternalName;

      field = mappingFields.find((f: any) => f.InternalName.toLowerCase().replace(/[^a-z0-9]/g, '') === normalizedSearch);
      if (field) return field.InternalName;

      return fallback;
    };

    const titleField = mappingFields.find((f: any) => f.InternalName === "Title");
    const titleFieldTitle = titleField ? (titleField.Title || "").toLowerCase().trim() : "";
    const isTitleEmploye = titleFieldTitle === "employe" || titleFieldTitle === "employee";

    let employeeNameFieldName = "Title";
    if (isTitleEmploye) {
      employeeNameFieldName = "Title";
    } else {
      const f = mappingFields.find((x: any) => {
        const title = (x.Title || "").toLowerCase().trim();
        const internal = (x.InternalName || "").toLowerCase().trim();
        return title === "employe" || title === "employee" || internal === "employe" || internal === "employee";
      });
      if (f) employeeNameFieldName = f.InternalName;
    }

    const employeeIdFieldName = (() => {
      const searchFields = mappingFields.filter((f: any) => f.InternalName !== employeeNameFieldName);

      const findFieldInList = (list: any[], search: string): string | null => {
        let f = list.find((x: any) => (x.Title || "").toLowerCase().trim() === search.toLowerCase());
        if (f) return f.InternalName;
        f = list.find((x: any) => x.InternalName.toLowerCase().trim() === search.toLowerCase());
        if (f) return f.InternalName;
        const norm = search.toLowerCase().replace(/[^a-z0-9]/g, '');
        f = list.find((x: any) => (x.Title || "").toLowerCase().replace(/[^a-z0-9]/g, '') === norm);
        if (f) return f.InternalName;
        return null;
      };

      return (
        findFieldInList(searchFields, "employe id") ||
        findFieldInList(searchFields, "employee id") ||
        findFieldInList(searchFields, "employeeid") ||
        findFieldInList(searchFields, "employeid") ||
        findFieldInList(searchFields, "employe") ||
        "EmployeeID"
      );
    })();

    const assetNameFieldName = findField("asset name", "AssetName");
    const serialNumberFieldName = findField("serial number", "SerialNumber");
    const priorityFieldName = findField("priority", "Priority");
    const requestedDateFieldName = findField("requested date", "RequestedDate");
    const reasonFieldName = findField("reason for request", "ReasonforRequest");
    const assignedDateFieldName = findField("assigned date", "AssignedDate");

    const payload: any = {};

    // Map Employee Name to the Employee column (either Title or a custom field)
    if (employeeNameFieldName === "Title") {
      payload["Title"] = employeeName;
    } else {
      payload["Title"] = `Assignment of ${assetName}`;

      const empFieldObj = mappingFields.find((f: any) => f.InternalName === employeeNameFieldName);
      const isEmpFieldPerson = empFieldObj && (empFieldObj.TypeAsString === "User" || empFieldObj.TypeAsString === "UserMulti");

      if (isEmpFieldPerson && assignedToId !== null) {
        payload[`${employeeNameFieldName}Id`] = assignedToId;
      } else {
        payload[employeeNameFieldName] = employeeName;
      }
    }

    // Map Employee ID to the Employee ID column
    const empIdFieldObj = mappingFields.find((f: any) => f.InternalName === employeeIdFieldName);
    const isEmpIdFieldPerson = empIdFieldObj && (empIdFieldObj.TypeAsString === "User" || empIdFieldObj.TypeAsString === "UserMulti");

    if (isEmpIdFieldPerson && assignedToId !== null) {
      payload[`${employeeIdFieldName}Id`] = assignedToId;
    } else {
      payload[employeeIdFieldName] = employeeId;
    }

    // Map other columns
    payload[assetNameFieldName] = assetName;
    payload[serialNumberFieldName] = serialNumber;
    payload[priorityFieldName] = priority;
    payload[requestedDateFieldName] = requestedDate;
    payload[reasonFieldName] = reason;
    payload[assignedDateFieldName] = assignedDate;

    return payload;
  }

  private static async _writeToMappingList(
    employeeName: string,
    employeeId: string,
    employeeEmail: string,
    assetName: string,
    serialNumber: string,
    priority: string,
    requestedDate: string,
    reason: string,
    assignedDate: string
  ): Promise<void> {
    const sp = getSP();
    const mappingList = await InventoryService.getMappingList();
    const mappingFields: any[] = await mappingList.fields.select("InternalName", "Title", "TypeAsString")();

    // Resolve user ID
    let assignedToId: number | null = null;
    if (employeeEmail) {
      try {
        const user: any = await sp.web.ensureUser(employeeEmail);
        assignedToId = user.data ? user.data.Id : user.Id;
      } catch (e) {
        console.warn(`Could not resolve user ${employeeEmail} in SharePoint during mapping write.`, e);
      }
    }

    const titleField = mappingFields.find((f: any) => f.InternalName === "Title");
    const titleFieldTitle = titleField ? (titleField.Title || "").toLowerCase().trim() : "";
    const isTitleEmploye = titleFieldTitle === "employe" || titleFieldTitle === "employee";

    const employeFieldName = isTitleEmploye ? "Title" : (() => {
      let f = mappingFields.find((x: any) => x.InternalName.toLowerCase() === "employe");
      if (f) return f.InternalName;
      f = mappingFields.find((x: any) => (x.Title || "").toLowerCase() === "employe");
      return f ? f.InternalName : "Employe";
    })();

    const employeFieldObj = mappingFields.find((f: any) => f.InternalName === employeFieldName);
    const isEmployePerson = employeFieldObj && (employeFieldObj.TypeAsString === "User" || employeFieldObj.TypeAsString === "UserMulti");

    // Dynamic field resolving
    const dynamicMappingPayload = InventoryService._resolveMappingPayload(
      mappingFields,
      employeeName,
      employeeId,
      assetName,
      serialNumber,
      priority,
      requestedDate,
      reason,
      assignedDate,
      assignedToId,
      isEmployePerson
    );

    const mappingPayloadsToTry: any[] = [];
    mappingPayloadsToTry.push(dynamicMappingPayload);

    // Fallbacks
    // Fallback 1: Title = employeeName, Employe = employeeId
    mappingPayloadsToTry.push({
      Title: employeeName,
      Employe: employeeId,
      AssetName: assetName,
      SerialNumber: serialNumber,
      Priority: priority,
      RequestedDate: requestedDate,
      ReasonforRequest: reason,
      AssignedDate: assignedDate
    });

    // Fallback 2: Title = employeeName, EmployeeID = employeeId
    mappingPayloadsToTry.push({
      Title: employeeName,
      EmployeeID: employeeId,
      AssetName: assetName,
      SerialNumber: serialNumber,
      Priority: priority,
      RequestedDate: requestedDate,
      ReasonforRequest: reason,
      AssignedDate: assignedDate
    });

    // Fallback 3: Title = employeeName, EmployeId = assignedToId, EmployeeID = employeeId
    if (assignedToId !== null) {
      mappingPayloadsToTry.push({
        Title: employeeName,
        EmployeId: assignedToId,
        EmployeeID: employeeId,
        AssetName: assetName,
        SerialNumber: serialNumber,
        Priority: priority,
        RequestedDate: requestedDate,
        ReasonforRequest: reason,
        AssignedDate: assignedDate
      });
    }

    // Fallback 4: Title = Assignment of Asset, Employe = employeeName, EmployeeID = employeeId
    mappingPayloadsToTry.push({
      Title: `Assignment of ${assetName}`,
      Employe: employeeName,
      EmployeeID: employeeId,
      AssetName: assetName,
      SerialNumber: serialNumber,
      Priority: priority,
      RequestedDate: requestedDate,
      ReasonforRequest: reason,
      AssignedDate: assignedDate
    });

    // Fallback 5: Title = Assignment of Asset, EmployeId = assignedToId, EmployeeID = employeeId
    if (assignedToId !== null) {
      mappingPayloadsToTry.push({
        Title: `Assignment of ${assetName}`,
        EmployeId: assignedToId,
        EmployeeID: employeeId,
        AssetName: assetName,
        SerialNumber: serialNumber,
        Priority: priority,
        RequestedDate: requestedDate,
        ReasonforRequest: reason,
        AssignedDate: assignedDate
      });
    }

    let mappingSuccess = false;
    let mappingLastErr: any;
    for (const p of mappingPayloadsToTry) {
      try {
        const filteredPayload = InventoryService._filterPayloadBySchema(p, mappingFields);
        console.log(`Attempting to write mapping payload for ${employeeName}:`, JSON.stringify(filteredPayload));
        await mappingList.items.add(filteredPayload);
        mappingSuccess = true;
        console.log(`Successfully added mapping record for ${employeeName}`);
        break;
      } catch (err) {
        mappingLastErr = err;
      }
    }

    if (!mappingSuccess) {
      console.warn(`Failed to write mapping record for ${employeeName} to Mapping List.`, mappingLastErr);
    }
  }

  private static _filterPayloadBySchema(payload: any, fields: any[]): any {
    const cleanPayload: any = {};
    const lowerFields = fields.map((f: any) => (f.InternalName || "").toLowerCase());
    
    // Always copy Title
    cleanPayload["Title"] = payload["Title"] || "";

    for (const key of Object.keys(payload)) {
      if (key === "Title") continue;
      
      const keyLower = key.toLowerCase();
      // Handle User ID fields (e.g., EmployeeId, EmployeId)
      if (keyLower.endsWith("id")) {
        const baseKey = key.substring(0, key.length - 2);
        const baseKeyLower = baseKey.toLowerCase();
        if (lowerFields.indexOf(baseKeyLower) >= 0 || lowerFields.indexOf(keyLower) >= 0) {
          cleanPayload[key] = payload[key];
          continue;
        }
      }
      
      if (lowerFields.indexOf(keyLower) >= 0) {
        cleanPayload[key] = payload[key];
      }
    }
    return cleanPayload;
  }

  public static async assignAssetsToEmployee(
    assetIds: string[],
    employeeName: string,
    employeeEmail: string,
    adminName: string,
    employeeId?: string
  ): Promise<void> {
    const sp = getSP();
    const list = await InventoryService.getInventoryList();
    let assignedToId: number | null = null;

    // Ensure 'Note' column exists to guarantee we have a place to save the Assignee
    try {
      const fields = await list.fields();
      if (!fields.some((f: any) => f.InternalName === 'Note')) {
        await list.fields.addMultilineText('Note', { NumberOfLines: 6, RichText: false });
        console.log("Automatically created 'Note' column in SharePoint list.");
      }
    } catch (e) {
      console.warn("Failed to check or create Note column", e);
    }

    // Try to resolve the user in SharePoint by email
    try {
      const user: any = await sp.web.ensureUser(employeeEmail);
      assignedToId = user.data ? user.data.Id : user.Id;
    } catch (e) {
      console.warn(`Could not resolve user ${employeeEmail} in SharePoint. Falling back to string assignment if column allows.`, e);
    }

    const updatePromises = assetIds.map(async (assetId) => {
      // 1. Get asset details first
      let assetItem: any = null;
      try {
        assetItem = await list.items.getById(parseInt(assetId))();
      } catch (e) {
        console.warn(`Could not fetch details for asset ${assetId}`, e);
      }

      const assetName = assetItem ? (assetItem.AssetName || assetItem.Asset_x0020_Name || assetItem.Asset || assetItem.Title || "") : "";
      const assetType = assetItem ? (assetItem.AssetType || assetItem.Asset_x0020_Type || assetItem.Type || "") : "";
      const serialNumber = assetItem ? (assetItem.SerialNumber || assetItem.Serial_x0020_Number || "") : "";

      // 2. Perform the update to InventoryList
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
      payloadsToTry.push({ ...baseStatus, Notes: `Assigned to: ${employeeName}` });
      payloadsToTry.push({ Status: `Assigned to: ${employeeName}`, AssetStatus: `Assigned to: ${employeeName}` }); // Fallback to Status column
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

      // 3. Find matching request in RequestList
      let priority = "Medium";
      let requestedDate = "";
      let reason = "Direct Assignment";
      let matchingRequest: any = null;

      try {
        const requests = await InventoryService.getRequests();
        matchingRequest = requests.find(r => {
          const isEmployeeMatch = (employeeId && r.employeeId && r.employeeId.toLowerCase() === employeeId.toLowerCase()) ||
            (employeeName && r.requesterName && r.requesterName.toLowerCase() === employeeName.toLowerCase());

          const isAssetMatch = assetType && r.assetTitle && r.assetTitle.toLowerCase() === assetType.toLowerCase();

          return isEmployeeMatch && isAssetMatch && r.status === 'Approved' && r.assetStatus === 'Pending';
        });

        if (matchingRequest) {
          priority = matchingRequest.priority || "Medium";
          requestedDate = matchingRequest.requestDate || "";
          reason = matchingRequest.reason || "Direct Assignment";
        }
      } catch (err) {
        console.warn("Failed to find matching approved request in RequestList", err);
      }

      // Format dates properly
      const formatDate = (dateStr?: string): string => {
        if (!dateStr) {
          const d = new Date();
          const day = ("0" + d.getDate()).slice(-2);
          const month = ("0" + (d.getMonth() + 1)).slice(-2);
          return `${day}/${month}/${d.getFullYear()}`;
        }
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
          return dateStr;
        }
        const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(dateStr);
        if (match) {
          return `${match[3]}/${match[2]}/${match[1]}`;
        }
        return dateStr;
      };

      const finalRequestedDate = formatDate(requestedDate);
      const finalAssignedDate = formatDate(new Date().toISOString());

      // 4. Update the matching request status to allocated (assetStatus = 'Approved')
      if (matchingRequest) {
        try {
          await InventoryService.updateAssetStatus(parseInt(matchingRequest.id, 10), 'Approved', adminName);
        } catch (err) {
          console.warn(`Failed to update assetStatus to Approved for Request ${matchingRequest.id}`, err);
        }
      }

      // 5. Update Mapping List
      try {
        await InventoryService._ensureMappingListFields();
        await InventoryService._writeToMappingList(
          employeeName,
          employeeId || "",
          employeeEmail,
          assetName,
          serialNumber,
          priority,
          finalRequestedDate,
          reason,
          finalAssignedDate
        );
      } catch (err) {
        console.warn("Failed to execute Mapping List update logic. Continuing.", err);
      }

      await this.addAuditLog({
        title: `Asset activated and assigned to ${employeeName}`,
        action: 'Activated',
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

  public static async syncExistingAssignmentsToMappingList(adminName: string): Promise<{ checkedCount: number; syncedCount: number }> {
    let checkedCount = 0;
    let syncedCount = 0;
    try {
      console.log("Starting Mapping List sync...");
      const list = await InventoryService.getInventoryList();
      const items = await InventoryService.getItems();

      // Filter for assigned items matching the 5 active employees
      const assignedItems = items.filter(item => {
        return EMPLOYEES.some(emp => {
          const normalize = (value: string | undefined): string => (value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          const activeUser = normalize(emp.name);
          if (!activeUser) return false;

          const assignedNorm = normalize(item.assignedTo);
          const isAssigned = assignedNorm && (assignedNorm === activeUser || assignedNorm.includes(activeUser) || activeUser.includes(assignedNorm));
          const isNoted = (item.note || '').toLowerCase().includes('assigned to:') && normalize(item.note).includes(activeUser);
          const isStatus = (item.status || '').toLowerCase().includes('assigned to:') && normalize(item.status).includes(activeUser);

          return !!(isAssigned || isNoted || isStatus);
        });
      });

      checkedCount = assignedItems.length;

      if (assignedItems.length === 0) {
        console.log("No assigned assets found for the 5 active employees.");
        return { checkedCount, syncedCount };
      }

      await InventoryService._ensureMappingListFields();
      const mappingList = await InventoryService.getMappingList();
      const mappingItems = await mappingList.items();

      const requests = await InventoryService.getRequests();

      // Resolve the Mapping List field names
      const mappingFields: any[] = await mappingList.fields.select("InternalName", "Title", "TypeAsString")();

      const findMappingField = (searchStr: string, fallback: string): string => {
        let field = mappingFields.find((f: any) => f.InternalName.toLowerCase() === searchStr.toLowerCase());
        if (field) return field.InternalName;

        field = mappingFields.find((f: any) => f.Title.toLowerCase() === searchStr.toLowerCase());
        if (field) return field.InternalName;

        const normalizedSearch = searchStr.toLowerCase().replace(/[^a-z0-9]/g, '');
        field = mappingFields.find((f: any) => f.Title.toLowerCase().replace(/[^a-z0-9]/g, '') === normalizedSearch);
        if (field) return field.InternalName;

        field = mappingFields.find((f: any) => f.Title.toLowerCase().replace(/[^a-z0-9]/g, '').indexOf(normalizedSearch) >= 0);
        if (field) return field.InternalName;

        return fallback;
      };

      const serialNumberFieldName = findMappingField("serialnumber", "SerialNumber");

      for (const asset of assignedItems) {
        // Check if this asset is already in the Mapping List
        const alreadyMapped = mappingItems.some((m: any) => {
          const mSerial = m[serialNumberFieldName] || m.SerialNumber || m.Serial_x0020_Number || "";
          return mSerial.toString().toLowerCase() === asset.serialNumber.toString().toLowerCase();
        });

        if (alreadyMapped) {
          console.log(`Asset ${asset.assetName || asset.title} (${asset.serialNumber}) is already in Mapping List.`);
          continue;
        }

        const employee = EMPLOYEES.find(emp => {
          const normalize = (value: string | undefined): string => (value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
          const activeUser = normalize(emp.name);
          if (!activeUser) return false;

          const assignedNorm = normalize(asset.assignedTo);
          const isAssigned = assignedNorm && (assignedNorm === activeUser || assignedNorm.includes(activeUser) || activeUser.includes(assignedNorm));
          const isNoted = (asset.note || '').toLowerCase().includes('assigned to:') && normalize(asset.note).includes(activeUser);
          const isStatus = (asset.status || '').toLowerCase().includes('assigned to:') && normalize(asset.status).includes(activeUser);

          return !!(isAssigned || isNoted || isStatus);
        });

        const employeeName = employee ? employee.name : (asset.assignedTo || "Unknown");
        const employeeId = employee ? employee.id : "";
        const employeeEmail = employee ? employee.email : "";

        const assetAssignedTo = employeeName;
        console.log(`Syncing missing assigned asset to Mapping List: ${asset.assetName || asset.title} (${asset.serialNumber}) for ${assetAssignedTo}`);

        // Find matching request
        let priority = "Medium";
        let requestedDate = "";
        let reason = "Direct Assignment";
        let matchingRequest: any = null;

        const assetType = asset.assetType || "";
        matchingRequest = requests.find(r => {
          const isEmployeeMatch = (employeeId && r.employeeId && r.employeeId.toLowerCase() === employeeId.toLowerCase()) ||
            (employeeName && r.requesterName && r.requesterName.toLowerCase() === employeeName.toLowerCase());

          const isAssetMatch = assetType && r.assetTitle && r.assetTitle.toLowerCase() === assetType.toLowerCase();

          return isEmployeeMatch && isAssetMatch && r.status === 'Approved';
        });

        if (matchingRequest) {
          priority = matchingRequest.priority || "Medium";
          requestedDate = matchingRequest.requestDate || "";
          reason = matchingRequest.reason || "Direct Assignment";
        }

        // Format dates properly
        const formatDate = (dateStr?: string): string => {
          if (!dateStr) {
            const d = new Date();
            const day = ("0" + d.getDate()).slice(-2);
            const month = ("0" + (d.getMonth() + 1)).slice(-2);
            return `${day}/${month}/${d.getFullYear()}`;
          }
          if (/^\d{2}\/\d{2}\/\d{4}$/.test(dateStr)) {
            return dateStr;
          }
          const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(dateStr);
          if (match) {
            return `${match[3]}/${match[2]}/${match[1]}`;
          }
          return dateStr;
        };

        const finalRequestedDate = formatDate(requestedDate);
        const finalAssignedDate = formatDate(asset.assignedDate || new Date().toISOString());

        // Update the matching request status to allocated (assetStatus = 'Approved')
        if (matchingRequest && matchingRequest.assetStatus === 'Pending') {
          try {
            await InventoryService.updateAssetStatus(parseInt(matchingRequest.id, 10), 'Approved', adminName);
          } catch (err) {
            console.warn(`Failed to update assetStatus to Approved for Request ${matchingRequest.id}`, err);
          }
        }

        // Add to Mapping List
        const assetName = asset.assetName || asset.title || "";
        const serialNumber = asset.serialNumber || "";

        try {
          await InventoryService._writeToMappingList(
            employeeName,
            employeeId,
            employeeEmail,
            assetName,
            serialNumber,
            priority,
            finalRequestedDate,
            reason,
            finalAssignedDate
          );
          syncedCount++;
        } catch (err) {
          console.warn(`Failed to write synced record for ${employeeName} to Mapping List.`, err);
        }
      }
    } catch (e) {
      console.error("Failed to sync existing assignments to Mapping List:", e);
      throw e;
    }
    return { checkedCount, syncedCount };
  }

  public static async diagnoseMappingListFields(): Promise<string> {
    try {
      const mappingList = await InventoryService.getMappingList();
      const fields: any[] = await mappingList.fields.select("InternalName", "Title", "TypeAsString")();
      const items = await mappingList.items.select("ID")();

      let output = `Mapping List Diagnostics:\n`;
      output += `- List URL Title: "${mappingList.Title || "Mapping List"}"\n`;
      output += `- Record Count: ${items.length}\n`;
      output += `- Available Fields in List:\n`;

      fields.forEach((f: any) => {
        output += `  * InternalName: "${f.InternalName}", Title: "${f.Title}", Type: "${f.TypeAsString}"\n`;
      });

      return output;
    } catch (e: any) {
      return `Failed to diagnose Mapping List schema. Error: ${e.message || JSON.stringify(e)}`;
    }
  }

  public static async getAuditLogs(): Promise<IEventLog[]> {
    let logs: IEventLog[] = [];
    const processedEventLogIds = new Set<string>();

    const formatTimestamp = (isoString?: string): string => {
      if (!isoString) return new Date().toISOString().replace('T', ' ').substring(0, 19);
      try {
        const formatted = isoString.replace('T', ' ');
        if (formatted.indexOf('.') > 0) {
          return formatted.substring(0, formatted.indexOf('.'));
        }
        if (formatted.indexOf('Z') > 0) {
          return formatted.substring(0, formatted.indexOf('Z'));
        }
        return formatted.substring(0, 19);
      } catch {
        return isoString.replace('T', ' ').substring(0, 19);
      }
    };

    const sp = getSP();

    // 1. Fetch from EventLogList (the audit logging list)
    try {
      const eventLogList = sp.web.lists.getByTitle(InventoryService.EVENT_LOG_LIST);
      const eventItems = await eventLogList.items.select("ID", "Title", "Action", "EntityType", "EntityId", "Details", "User", "Created")();

      eventItems.forEach((item: any) => {
        let assetName = "";
        let detailsText = item.Details || "";
        let actionText = item.Action || "";

        // Parse JSON details if possible
        if (detailsText.trim().startsWith("{") && detailsText.trim().endsWith("}")) {
          try {
            const parsed = JSON.parse(detailsText);
            assetName = parsed.assetName || parsed.assetTitle || parsed.Title || "";
            if (parsed.lifecycle) {
              if (parsed.lifecycle === "DirectAssignment") {
                detailsText = `Asset assigned to employee`;
                actionText = "admin assigned";
              } else if (parsed.lifecycle === "Approved") {
                detailsText = `Approved request for ${assetName || "Asset"}`;
                actionText = "manager approved";
              } else if (parsed.lifecycle === "Rejected" || parsed.lifecycle === "Declined") {
                detailsText = `Rejected request for ${assetName || "Asset"}`;
                actionText = "manager rejected";
              } else if (parsed.lifecycle === "Submitted") {
                detailsText = `Submitted asset request for ${assetName || "Asset"}`;
                actionText = "created";
              } else if (parsed.lifecycle === "ReturnRequested") {
                detailsText = `Requested return of ${assetName || "Asset"}`;
                actionText = "return requested";
              } else if (parsed.lifecycle === "ReturnApproved") {
                detailsText = `Approved return request for ${assetName || "Asset"}`;
                actionText = "return approved";
              } else if (parsed.lifecycle === "ReturnRejected") {
                detailsText = `Rejected return request for ${assetName || "Asset"}`;
                actionText = "return rejected";
              } else if (parsed.lifecycle === "ReturnCompleted") {
                detailsText = `Completed return of ${assetName || "Asset"}`;
                actionText = "return completed";
              }
            } else if (parsed.assetStatus === "Approved" || parsed.lifecycle === "AssetStatusUpdated") {
              detailsText = `Assigned ${assetName || "Asset"}`;
              actionText = "admin assigned";
            }
          } catch (e) {
            // Keep detailsText as string
          }
        }

        const titleLower = (item.Title || "").toLowerCase();
        const actionLower = (item.Action || "").toLowerCase();
        const detailsLower = (detailsText || "").toLowerCase();

        // High fidelity text mapping
        if (actionLower === "create") {
          actionText = "created";
        } else if (actionLower === "delete") {
          actionText = "deleted";
        }

        if (titleLower.includes("directly assigned") || detailsLower.includes("assigned to employee") || detailsLower.includes("assigned to:")) {
          actionText = "admin assigned";
          detailsText = "Asset assigned to employee";
        } else if (titleLower.includes("approved return request") || detailsLower.includes("approved return request")) {
          actionText = "return approved";
          detailsText = `Approved return request for ${assetName || "Asset"}`;
        } else if (titleLower.includes("rejected return request") || detailsLower.includes("rejected return request")) {
          actionText = "return rejected";
          detailsText = `Rejected return request for ${assetName || "Asset"}`;
        } else if (titleLower.includes("completed return") || detailsLower.includes("completed return")) {
          actionText = "return completed";
          detailsText = `Completed return of ${assetName || "Asset"}`;
        } else if (titleLower.includes("requested return") || detailsLower.includes("requested return")) {
          actionText = "return requested";
          detailsText = `Requested return of ${assetName || "Asset"}`;
        } else if (titleLower.includes("approved request") || detailsLower.includes("approved request")) {
          actionText = "manager approved";
          detailsText = `Approved request for ${assetName || "Asset"}`;
        } else if (titleLower.includes("rejected request") || detailsLower.includes("rejected request") || titleLower.includes("declined request") || detailsLower.includes("declined request")) {
          actionText = "manager rejected";
          detailsText = `Rejected request for ${assetName || "Asset"}`;
        } else if (titleLower.includes("created request") || detailsLower.includes("submitted asset request")) {
          actionText = "created";
          detailsText = `Submitted asset request for ${assetName || "Asset"}`;
        } else if (titleLower.includes("created asset") || detailsLower.includes("added to inventory")) {
          actionText = "created";
          detailsText = "Asset was added to inventory";
        } else if (titleLower.includes("deleted asset") || detailsLower.includes("retired due to damage")) {
          actionText = "deleted";
          detailsText = "Asset retired due to damage";
        } else if (detailsLower.includes("under investigation") || detailsLower.includes("in progress")) {
          actionText = "status updated to in progress";
          detailsText = "Under investigation";
        } else if (detailsLower.includes("resolved") || detailsLower.includes("replacement cable")) {
          actionText = "status updated to resolved";
          detailsText = "Replacement cable sourced";
        }

        // Try to extract assetName if empty
        if (!assetName) {
          if (titleLower.includes("created asset:")) {
            assetName = item.Title.substring(titleLower.indexOf("created asset:") + 14).trim();
          } else if (titleLower.includes("updated asset:")) {
            assetName = item.Title.substring(titleLower.indexOf("updated asset:") + 14).trim();
          } else if (titleLower.includes("deleted asset:")) {
            assetName = item.Title.substring(titleLower.indexOf("deleted asset:") + 14).trim();
          } else if (titleLower.includes("request for")) {
            assetName = item.Title.substring(titleLower.indexOf("request for") + 11).trim();
          }
        }

        logs.push({
          id: `event-log-${item.ID}`,
          title: item.Title || "",
          action: actionText,
          entityType: item.EntityType || "",
          entityId: item.EntityId || "",
          assetName: assetName,
          details: detailsText,
          user: item.User || "System",
          timestamp: formatTimestamp(item.Created)
        });

        processedEventLogIds.add(`${item.EntityType}-${item.EntityId}-${item.Action}-${formatTimestamp(item.Created).substring(0, 16)}`);
      });
    } catch (err) {
      console.warn("Could not fetch EventLogList, relying on live list generation", err);
    }

    // 2. Fetch from InventoryList for dynamic fallback logs
    try {
      const list = await InventoryService.getInventoryList();
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString")();
      const findFieldInternalName = (searchStr: string, fallback: string): string => {
        let field = fields.find((f: any) => f.InternalName.toLowerCase() === searchStr.toLowerCase());
        if (field) return field.InternalName;
        field = fields.find((f: any) => f.InternalName.toLowerCase().replace(/_x0020_/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        if (field) return field.InternalName;
        field = fields.find((f: any) => f.Title.toLowerCase().replace(/[^a-z0-9]/g, '').indexOf(searchStr.toLowerCase()) >= 0);
        return field ? field.InternalName : fallback;
      };

      const assetNameKey = findFieldInternalName("assetname", "AssetName");
      const statusKey = findFieldInternalName("status", "Status");

      const inventoryItems = await list.items.select("ID", "Title", assetNameKey, statusKey, "Created", "Modified", "Author/Title", "Editor/Title").expand("Author", "Editor")();

      inventoryItems.forEach((item: any) => {
        const itemCreated = formatTimestamp(item.Created);
        const itemModified = formatTimestamp(item.Modified);
        const assetName = item[assetNameKey] || item.Title || "";
        const itemStatus = item[statusKey] || "";

        // Add dynamic Create event
        const createKey = `Asset-${item.ID}-Create-${itemCreated.substring(0, 16)}`;
        if (!processedEventLogIds.has(createKey)) {
          logs.push({
            id: `asset-create-${item.ID}`,
            title: `Created Asset: ${item.Title}`,
            action: 'created',
            entityType: 'Asset',
            entityId: item.ID.toString(),
            assetName: assetName,
            details: `Asset was added to inventory`,
            user: item.Author?.Title || "System",
            timestamp: itemCreated
          });
        }

        // Add dynamic Update event if item has been updated
        if (item.Modified && item.Created && new Date(item.Modified).getTime() - new Date(item.Created).getTime() > 5000) {
          const updateKey = `Asset-${item.ID}-Update-${itemModified.substring(0, 16)}`;
          const updateKeyAssigned = `Asset-${item.ID}-admin assigned-${itemModified.substring(0, 16)}`;
          if (!processedEventLogIds.has(updateKey) && !processedEventLogIds.has(updateKeyAssigned)) {
            let actionText = "updated";
            let detailsText = "Asset details were modified";

            if (itemStatus.toLowerCase().includes("assigned") || itemStatus.toLowerCase() === "active" || itemStatus.toLowerCase() === "in use") {
              actionText = "admin assigned";
              detailsText = "Asset assigned to employee";
            } else if (itemStatus.toLowerCase().includes("maintenance")) {
              actionText = "status updated to in progress";
              detailsText = "Under investigation";
            } else if (itemStatus.toLowerCase() === "in stock" || itemStatus.toLowerCase() === "yes") {
              actionText = "status updated to resolved";
              detailsText = "Replacement cable sourced";
            }

            logs.push({
              id: `asset-update-${item.ID}-${new Date(item.Modified).getTime()}`,
              title: `Updated Asset: ${item.Title}`,
              action: actionText,
              entityType: 'Asset',
              entityId: item.ID.toString(),
              assetName: assetName,
              details: detailsText,
              user: item.Editor?.Title || "System",
              timestamp: itemModified
            });
          }
        }
      });
    } catch (err) {
      console.warn("Could not fetch InventoryList for audit logs", err);
    }

    // 3. Fetch from RequestList for dynamic fallback logs
    try {
      const reqList = await InventoryService.getRequestList();
      const requestItems = await this._fetchItemsWithExpandedUsers(reqList);

      requestItems.forEach((item: any) => {
        const keys = Object.keys(item);
        const findKey = (searchStr: string) => {
          const nonIdMatch = keys.find(k => {
            const kl = k.toLowerCase().replace(/_x0020_/g, '');
            return kl.indexOf(searchStr) >= 0 && !kl.endsWith("id");
          });
          if (nonIdMatch) return nonIdMatch;
          return keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').indexOf(searchStr) >= 0);
        };

        const employeeKey = findKey("requester") || findKey("employee") || "Employee";
        const selectAssetKey = findKey("assettype") || findKey("selectasset") || findKey("type") || "SelectAsset";
        const statusKey = keys.find(key => this._isBusinessStatusKey(key)) || "RequestStatus";
        const assetStatusKey = findKey("assetstatus") || "AssetStatus";

        const reqAssetName = item[selectAssetKey] || item.Title || "Unknown Asset";
        const rawEmp = item[employeeKey] || item.Employee || item.Author;
        const reqUser = (() => {
          if (!rawEmp) return item.Title || "System";
          if (typeof rawEmp === 'string') return rawEmp;
          if (Array.isArray(rawEmp)) return rawEmp.map((a: any) => a.Title || a.Name || "").join(', ');
          if (typeof rawEmp === 'object') return rawEmp.Title || rawEmp.Name || JSON.stringify(rawEmp);
          return rawEmp.toString();
        })();

        const itemCreated = formatTimestamp(item.Created);
        const itemModified = formatTimestamp(item.Modified);
        const requestStatus = item[statusKey] || "";
        const assetStatus = item[assetStatusKey] || "";

        // Add dynamic Create event
        const createKey = `Request-${item.ID}-Create-${itemCreated.substring(0, 16)}`;
        if (!processedEventLogIds.has(createKey)) {
          logs.push({
            id: `request-create-${item.ID}`,
            title: `Created Request: ${reqAssetName}`,
            action: 'created',
            entityType: 'Request',
            entityId: item.ID.toString(),
            assetName: reqAssetName,
            details: `Submitted asset request for ${reqAssetName}`,
            user: reqUser,
            timestamp: itemCreated
          });
        }

        // Add dynamic Update event
        if (item.Modified && item.Created && new Date(item.Modified).getTime() - new Date(item.Created).getTime() > 5000) {
          const updateKey = `Request-${item.ID}-Update-${itemModified.substring(0, 16)}`;
          if (!processedEventLogIds.has(updateKey)) {
            if (requestStatus.toLowerCase().includes("approv")) {
              logs.push({
                id: `request-approve-${item.ID}-${new Date(item.Modified).getTime()}`,
                title: `Approved Request: ${reqAssetName}`,
                action: 'manager approved',
                entityType: 'Request',
                entityId: item.ID.toString(),
                assetName: reqAssetName,
                details: `Approved request for ${reqAssetName}`,
                user: reqUser,
                timestamp: itemModified
              });

              if (assetStatus.toLowerCase().includes("approv")) {
                logs.push({
                  id: `request-assign-${item.ID}-${new Date(item.Modified).getTime()}`,
                  title: `Assigned Request: ${reqAssetName}`,
                  action: 'admin assigned',
                  entityType: 'Request',
                  entityId: item.ID.toString(),
                  assetName: reqAssetName,
                  details: `Assigned ${reqAssetName}`,
                  user: reqUser,
                  timestamp: itemModified
                });
              }
            } else if (requestStatus.toLowerCase().includes("reject") || requestStatus.toLowerCase().includes("declin")) {
              logs.push({
                id: `request-reject-${item.ID}-${new Date(item.Modified).getTime()}`,
                title: `Rejected Request: ${reqAssetName}`,
                action: 'manager rejected',
                entityType: 'Request',
                entityId: item.ID.toString(),
                assetName: reqAssetName,
                details: `Rejected request for ${reqAssetName}`,
                user: reqUser,
                timestamp: itemModified
              });
            } else {
              logs.push({
                id: `request-update-${item.ID}-${new Date(item.Modified).getTime()}`,
                title: `Updated Request: ${reqAssetName}`,
                action: 'updated',
                entityType: 'Request',
                entityId: item.ID.toString(),
                assetName: reqAssetName,
                details: `Request details were modified`,
                user: reqUser,
                timestamp: itemModified
              });
            }
          }
        }
      });
    } catch (err) {
      console.warn("Could not fetch RequestList for audit logs", err);
    }

    // Sort logs by timestamp descending
    logs.sort((a, b) => new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime());

    return logs;
  }

  public static async getRequestHistoryById(
    requestLookupId: string
  ): Promise<{ request: IRequest; lifecycle: IEventLog[] }> {
    await this._ensureRequestWorkflowFields();

    const normalizedRequestKey = this._normalizeRequestKey(requestLookupId);
    if (!normalizedRequestKey) {
      throw new Error("Request ID is required.");
    }

    const reqList = await InventoryService.getRequestList();
    let requestItems: any[] = [];
    try {
      const fields: any[] = await reqList.fields.select("InternalName", "Title")();
      const resolvedKeyName = InventoryService._resolveRequestKeyInternalName(fields);
      requestItems = await reqList
        .items.select("*")
        .filter(`${resolvedKeyName} eq '${normalizedRequestKey.replace(/'/g, "''")}'`)();
    } catch (filterError) {
      console.warn("RequestKey filter failed. Falling back to item ID based lookup.", filterError);
    }

    if (!requestItems.length) {
      const derivedIdMatch = /^REQ-(\d{1,})$/.exec(normalizedRequestKey.replace(/^REQ-0*/, "REQ-"));
      const parsedId = derivedIdMatch ? parseInt(derivedIdMatch[1], 10) : NaN;
      if (!Number.isNaN(parsedId)) {
        try {
          const requestById = await reqList
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

  private static async _fetchItemsWithExpandedUsers(list: any): Promise<any[]> {
    try {
      const fields = await list.fields.select("InternalName", "Title", "TypeAsString")();
      const selectFields = ["*"];
      const expandFields: string[] = [];

      fields.forEach((field: any) => {
        const name = field.InternalName;
        const type = field.TypeAsString;
        if (type === "User" || type === "UserMulti") {
          const nameLower = (name || '').toLowerCase();
          const titleLower = (field.Title || '').toLowerCase();
          if (
            nameLower.indexOf("employee") >= 0 ||
            nameLower.indexOf("requester") >= 0 ||
            nameLower.indexOf("assigned") >= 0 ||
            nameLower.indexOf("author") >= 0 ||
            nameLower.indexOf("editor") >= 0 ||
            titleLower.indexOf("employee") >= 0 ||
            titleLower.indexOf("requester") >= 0 ||
            titleLower.indexOf("assigned") >= 0
          ) {
            selectFields.push(`${name}/Title`, `${name}/Id`);
            expandFields.push(name);
          }
        }
      });

      if (expandFields.indexOf("Author") === -1) {
        selectFields.push("Author/Title", "Author/Id");
        expandFields.push("Author");
      }
      if (expandFields.indexOf("Editor") === -1) {
        selectFields.push("Editor/Title", "Editor/Id");
        expandFields.push("Editor");
      }

      return await list.items.select(selectFields.join(",")).expand(...expandFields)();
    } catch (error: any) {
      console.error("Dynamic user expansion failed:", error);
      throw new Error(`Dynamic user expansion failed: ${error.message || JSON.stringify(error)}`);
    }
  }

  public static async getReturnRequests(): Promise<IReturnRequest[]> {
    const sp = getSP();
    try {
      const list = sp.web.lists.getByTitle("ReturnRequestList");
      const items = await list.items.select("ID", "Title", "AssetID", "AssetName", "SerialNumber", "RequesterName", "RequesterEmail", "RequestDate", "ReturnReason", "ProposedCondition", "Status", "ManagerComment", "CompletedDate")();
      return items.map((item: any) => ({
        id: item.ID.toString(),
        title: item.Title || "",
        assetId: item.AssetID || item.AssetId || "",
        assetName: item.AssetName || "",
        serialNumber: item.SerialNumber || "",
        requesterName: item.RequesterName || "",
        requesterEmail: item.RequesterEmail || "",
        requestDate: item.RequestDate || "",
        returnReason: item.ReturnReason || "",
        proposedCondition: item.ProposedCondition || "",
        status: item.Status || "Pending",
        managerComment: item.ManagerComment || "",
        completedDate: item.CompletedDate || ""
      }));
    } catch (error) {
      console.warn("Could not fetch ReturnRequestList from SharePoint, using local storage fallback", error);
      try {
        const local = localStorage.getItem("inventory_return_requests");
        return local ? JSON.parse(local) : [];
      } catch {
        return [];
      }
    }
  }

  public static async addReturnRequest(request: Omit<IReturnRequest, 'id' | 'status'>, userDisplayName: string): Promise<void> {
    const sp = getSP();
    const newRequest: IReturnRequest = {
      ...request,
      id: `RR-${Date.now()}`,
      status: 'Pending'
    };

    try {
      const list = sp.web.lists.getByTitle("ReturnRequestList");
      await list.items.add({
        Title: request.title,
        AssetID: request.assetId,
        AssetName: request.assetName,
        SerialNumber: request.serialNumber,
        RequesterName: request.requesterName,
        RequesterEmail: request.requesterEmail || "",
        RequestDate: request.requestDate,
        ReturnReason: request.returnReason,
        ProposedCondition: request.proposedCondition,
        Status: "Pending"
      });

    } catch (error) {
      console.warn("Failed to save return request to ReturnRequestList in SharePoint, saving to local storage fallback", error);
      try {
        const local = localStorage.getItem("inventory_return_requests");
        const list: IReturnRequest[] = local ? JSON.parse(local) : [];
        list.push(newRequest);
        localStorage.setItem("inventory_return_requests", JSON.stringify(list));
      } catch (e) {
        console.error("Local storage save failed", e);
      }
    }

    try {
      const list = await InventoryService.getInventoryList();
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString")();
      const statusField = fields.find((f: any) => f.InternalName.toLowerCase() === "status" || f.Title.toLowerCase() === "status");
      const statusKey = statusField ? statusField.InternalName : "Status";
      
      await list.items.getById(parseInt(request.assetId)).update({
        [statusKey]: "Pending Return"
      });
    } catch (error) {
      console.warn("Failed to update asset status in SharePoint, doing local storage fallback for asset state", error);
    }

    try {
      await this.addAuditLog({
        title: `Requested Return & Deactivated: ${request.assetName}`,
        action: 'Deactivated',
        entityType: 'Asset',
        entityId: request.assetId,
        details: JSON.stringify({
          lifecycle: "ReturnRequested",
          assetId: request.assetId,
          assetName: request.assetName,
          serialNumber: request.serialNumber,
          requesterName: request.requesterName,
          returnReason: request.returnReason,
          proposedCondition: request.proposedCondition,
          requestedAt: new Date().toISOString()
        }),
        user: userDisplayName
      });
    } catch (e) {
      console.warn("Failed to add audit log for return request", e);
    }
  }

  public static async updateReturnRequestStatus(
    requestId: string,
    status: 'Approved' | 'Rejected' | 'Completed',
    managerComment: string,
    approverName: string,
    finalCondition?: string
  ): Promise<void> {
    const sp = getSP();

    const requests = await this.getReturnRequests();
    const req = requests.find(r => r.id === requestId);
    if (!req) {
      throw new Error(`Return request with ID ${requestId} not found.`);
    }

    let updatedSharePoint = false;
    if (requestId.indexOf("RR-") !== 0) {
      try {
        const list = sp.web.lists.getByTitle("ReturnRequestList");
        const payload: any = {
          Status: status,
          ManagerComment: managerComment
        };
        if (status === 'Completed') {
          payload.CompletedDate = new Date().toISOString().split('T')[0];
        }
        await list.items.getById(parseInt(requestId)).update(payload);
        updatedSharePoint = true;
      } catch (err) {
        console.warn("Failed to update return request status in SharePoint", err);
      }
    }

    if (!updatedSharePoint) {
      try {
        const local = localStorage.getItem("inventory_return_requests");
        if (local) {
          const list: IReturnRequest[] = JSON.parse(local);
          const updated = list.map(r => {
            if (r.id === requestId) {
              const updatedReq = { ...r, status, managerComment };
              if (status === 'Completed') {
                updatedReq.completedDate = new Date().toISOString().split('T')[0];
              }
              return updatedReq;
            }
            return r;
          });
          localStorage.setItem("inventory_return_requests", JSON.stringify(updated));
        }
      } catch (e) {
        console.error("Local storage update failed", e);
      }
    }

    try {
      const list = await InventoryService.getInventoryList();
      const fields: any[] = await list.fields.select("InternalName", "Title", "TypeAsString")();
      
      const findField = (searchStr: string, fallback: string): string => {
        const field = fields.find((f: any) => f.InternalName.toLowerCase() === searchStr.toLowerCase() || f.Title.toLowerCase() === searchStr.toLowerCase());
        return field ? field.InternalName : fallback;
      };

      const statusKey = findField("status", "Status");
      const assignedToKey = findField("assignedto", "AssignedTo");
      const conditionKey = findField("condition", "Condition");
      const noteKey = findField("note", "Note");

      const assetIdNum = parseInt(req.assetId, 10);

      if (status === 'Completed') {
        const condition = finalCondition || req.proposedCondition || "Good";
        let nextStatus = "In Stock";
        if (condition === "Poor" || condition === "Damaged") {
          nextStatus = "Under Maintenance";
        }

        const payload: any = {
          [statusKey]: nextStatus,
          [assignedToKey]: null,
          [conditionKey]: condition,
          [noteKey]: `Returned by employee. Manager Note: ${managerComment}`
        };

        const assignedFieldObj = fields.find(f => f.InternalName === assignedToKey);
        const isUserField = assignedFieldObj && (assignedFieldObj.TypeAsString === "User" || assignedFieldObj.TypeAsString === "UserMulti");
        if (isUserField) {
          payload[`${assignedToKey}Id`] = null;
        }

        try {
          await list.items.getById(assetIdNum).update(payload);
        } catch (err) {
          console.warn("SharePoint Asset update failed on Return Completion, doing fallback", err);
        }
      } else if (status === 'Approved') {
        try {
          await list.items.getById(assetIdNum).update({
            [statusKey]: "Return Approved"
          });
        } catch (err) {
          console.warn("SharePoint Asset update failed on Return Approval", err);
        }
      } else if (status === 'Rejected') {
        try {
          await list.items.getById(assetIdNum).update({
            [statusKey]: "Assigned"
          });
        } catch (err) {
          console.warn("SharePoint Asset update failed on Return Rejection", err);
        }
      }
    } catch (error) {
      console.warn("Failed to execute Asset inventory sync for return requests. Continuing.", error);
    }

    try {
      let logTitle = "";
      let lifecycle = "";
      if (status === 'Approved') {
        logTitle = `Approved Return Request for Asset: ${req.assetName}`;
        lifecycle = "ReturnApproved";
      } else if (status === 'Rejected') {
        logTitle = `Rejected Return Request for Asset: ${req.assetName}`;
        lifecycle = "ReturnRejected";
      } else if (status === 'Completed') {
        logTitle = `Completed Return for Asset: ${req.assetName}`;
        lifecycle = "ReturnCompleted";
      }

      let finalAction = 'Update';
      let finalTitle = logTitle;
      if (status === 'Approved') {
        finalAction = 'Deactivated';
        finalTitle = `Approved Return Request & Deactivated: ${req.assetName}`;
      } else if (status === 'Rejected') {
        finalAction = 'Activated';
        finalTitle = `Rejected Return Request & Reactivated: ${req.assetName}`;
      } else if (status === 'Completed') {
        const condition = finalCondition || req.proposedCondition || "Good";
        let nextStatus = "In Stock";
        if (condition === "Poor" || condition === "Damaged") {
          nextStatus = "Under Maintenance";
        }
        if (nextStatus === 'In Stock') {
          finalAction = 'Inactivated';
          finalTitle = `Completed Return & Inactivated: ${req.assetName} (Returned to Stock)`;
        } else {
          finalAction = 'Deactivated';
          finalTitle = `Completed Return & Deactivated: ${req.assetName} (Under Maintenance)`;
        }
      }

      await this.addAuditLog({
        title: finalTitle,
        action: finalAction,
        entityType: 'Asset',
        entityId: req.assetId,
        details: JSON.stringify({
          requestKey: req.id,
          lifecycle,
          assetName: req.assetName,
          requesterName: req.requesterName,
          changedBy: approverName,
          changedAt: new Date().toISOString(),
          managerComment,
          condition: finalCondition || req.proposedCondition
        }),
        user: approverName
      });
    } catch (e) {
      console.warn("Failed to add audit log for return request status update", e);
    }
  }
}
