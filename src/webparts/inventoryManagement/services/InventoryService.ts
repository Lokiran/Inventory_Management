import { IInventoryItem } from "../models/IInventoryItem";
import { IRequest } from "../models/IRequest";
import { IEventLog } from "../models/IEventLog";
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

      return items.map((item: any) => ({
        id: item.ID.toString(),
        title: item.Title || "",
        assetName: item[assetNameKey] || item.AssetName || item.Asset_x0020_Name || item.Asset || "",
        assetType: item[assetTypeKey] || item.AssetType || item.Asset_x0020_Type || item.Type || "",
        serialNumber: item[serialNumberKey] || item.SerialNumber || item.Serial_x0020_Number || "",
        purchaseDate: item[purchaseDateKey] || item.PurchaseDate || item.Purchase_x0020_Date || "",
        vendor: item[vendorKey] || item.Vendor || item.VendorName || "",
        condition: item[conditionKey] || item.Condition || item.AssetCondition || "",
        status: item[statusKey] || item.Status || item.AssetStatus || "",
        assignedTo: (() => {
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
        })(),
        assignedDate: item.Modified || "",
        warrantyExpiry: item[warrantyExpiryKey] || item.WarrantyExpiry || item.Warranty_x0020_Expiry || "",
        specifications: item[specificationsKey] || item.Specifications || item.SpecificationsText || item.Note || item.Notes || "",
        note: item[noteKey] || item.Note || item.Notes || ""
      }));
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
        title: `Created Asset: ${item.title}`,
        action: 'Create',
        entityType: 'Asset',
        entityId,
        details: JSON.stringify(item),
        user: userDisplayName
      });
    } catch (auditError) {
      console.warn("Failed to write audit log for newly created asset:", auditError);
    }
  }

  public static async addRequest(request: Omit<IRequest, 'id' | 'requestKey' | 'status' | 'requestDate'> & { status?: string }, userDisplayName: string = "Unknown"): Promise<void> {
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
          RequestStatus: initialStatus
        },
        {
          Title: `Request for ${request.assetTitle}`,
          EmployeeId: requesterId,
          EmployeeID: (request as any).employeeId || "",
          Priority: (request as any).priority || "Medium",
          Assettype: request.assetTitle,
          Quantity: request.quantity,
          ReasonforRequest: request.reason || "",
          RequestStatus: initialStatus
        },
        {
          Title: `Request for ${request.assetTitle}`,
          RequesterId: requesterId,
          EmployeeID: (request as any).employeeId || "",
          Priority: (request as any).priority || "Medium",
          Asset_x0020_type: request.assetTitle,
          Quantity: request.quantity,
          Reason_x0020_for_x0020_Request: request.reason || "",
          RequestStatus: initialStatus
        },
        {
          Title: `Request for ${request.assetTitle}`,
          RequesterId: requesterId,
          EmployeeID: (request as any).employeeId || "",
          Priority: (request as any).priority || "Medium",
          Assettype: request.assetTitle,
          Quantity: request.quantity,
          ReasonforRequest: request.reason || "",
          RequestStatus: initialStatus
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
        RequestStatus: initialStatus
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Assettype: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: initialStatus
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Requester: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Asset_x0020_type: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
        RequestStatus: initialStatus
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Requester: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Assettype: request.assetTitle,
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: initialStatus
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
        RequestStatus: initialStatus
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Select_x0020_Asset: request.assetTitle,
        Quantity: request.quantity,
        Reason_x0020_for_x0020_Request: request.reason || "",
        RequestStatus: initialStatus
      },
      {
        Title: `Request for ${request.assetTitle}`,
        Employee: request.requesterName,
        EmployeeID: (request as any).employeeId || "",
        Priority: (request as any).priority || "Medium",
        Quantity: request.quantity,
        ReasonforRequest: request.reason || "",
        RequestStatus: initialStatus
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
          await requestListInstance.items.getById(requestItemId)
            .update({
              [InventoryService.REQUEST_KEY_INTERNAL_NAME]: requestKey,
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

      return items.map((item: any) => {
        const rawStatus = item[statusKey] || item.Status || 'Pending';
        const normalizedStatus = (rawStatus || '').toString().toLowerCase();
        const status =
          (normalizedStatus.includes('approv')) ? 'Approved' :
          (normalizedStatus.includes('declin') || normalizedStatus.includes('reject')) ? 'Declined' :
          'Pending';
        const requestKey = this._extractRequestKey(item);

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
        const mappingList = await InventoryService.getMappingList();
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

        const employeFieldName = findMappingField("employe", "Employe");
        const employeeIdFieldName = findMappingField("employeid", "") || findMappingField("employeeid", "EmployeeID");
        const assetNameFieldName = findMappingField("assetname", "AssetName");
        const serialNumberFieldName = findMappingField("serialnumber", "SerialNumber");
        const priorityFieldName = findMappingField("priority", "Priority");
        const requestedDateFieldName = findMappingField("requesteddate", "RequestedDate");
        const reasonFieldName = findMappingField("reasonforrequest", "ReasonforRequest");
        const assignedDateFieldName = findMappingField("assigneddate", "AssignedDate");

        const employeFieldObj = mappingFields.find((f: any) => f.InternalName === employeFieldName);
        const isEmployePerson = employeFieldObj && (employeFieldObj.TypeAsString === "User" || employeFieldObj.TypeAsString === "UserMulti");

        const dynamicMappingPayload: any = {};
        if (employeFieldName === "Title") {
          dynamicMappingPayload["Title"] = employeeName;
        } else {
          dynamicMappingPayload["Title"] = `Assignment of ${assetName}`;
          if (isEmployePerson && assignedToId !== null) {
            dynamicMappingPayload[`${employeFieldName}Id`] = assignedToId;
          } else {
            dynamicMappingPayload[employeFieldName] = employeeName;
          }
        }

        dynamicMappingPayload[employeeIdFieldName] = employeeId || "";
        dynamicMappingPayload[assetNameFieldName] = assetName;
        dynamicMappingPayload[serialNumberFieldName] = serialNumber;
        dynamicMappingPayload[priorityFieldName] = priority;
        dynamicMappingPayload[requestedDateFieldName] = finalRequestedDate;
        dynamicMappingPayload[reasonFieldName] = reason;
        dynamicMappingPayload[assignedDateFieldName] = finalAssignedDate;

        const mappingPayloadsToTry: any[] = [];
        mappingPayloadsToTry.push(dynamicMappingPayload);

        if (assignedToId !== null) {
          mappingPayloadsToTry.push({
            Title: employeeName,
            EmployeId: assignedToId,
            EmployeeID: employeeId || "",
            AssetName: assetName,
            SerialNumber: serialNumber,
            Priority: priority,
            RequestedDate: finalRequestedDate,
            ReasonforRequest: reason,
            AssignedDate: finalAssignedDate
          });
          mappingPayloadsToTry.push({
            Title: `Assignment of ${assetName}`,
            EmployeId: assignedToId,
            EmployeeID: employeeId || "",
            AssetName: assetName,
            SerialNumber: serialNumber,
            Priority: priority,
            RequestedDate: finalRequestedDate,
            ReasonforRequest: reason,
            AssignedDate: finalAssignedDate
          });
          mappingPayloadsToTry.push({
            Title: `Assignment of ${assetName}`,
            EmployeId: assignedToId,
            Employee_x0020_ID: employeeId || "",
            Asset_x0020_Name: assetName,
            Serial_x0020_Number: serialNumber,
            Priority: priority,
            Requested_x0020_Date: finalRequestedDate,
            Reason_x0020_for_x0020_Request: reason,
            AssignedDate: finalAssignedDate
          });
        }

        mappingPayloadsToTry.push({
          Title: employeeName,
          Employe: employeeName,
          EmployeeID: employeeId || "",
          AssetName: assetName,
          SerialNumber: serialNumber,
          Priority: priority,
          RequestedDate: finalRequestedDate,
          ReasonforRequest: reason,
          AssignedDate: finalAssignedDate
        });
        mappingPayloadsToTry.push({
          Title: `Assignment of ${assetName}`,
          Employe: employeeName,
          EmployeeID: employeeId || "",
          AssetName: assetName,
          SerialNumber: serialNumber,
          Priority: priority,
          RequestedDate: finalRequestedDate,
          ReasonforRequest: reason,
          AssignedDate: finalAssignedDate
        });

        let mappingSuccess = false;
        let mappingLastErr: any;
        for (const p of mappingPayloadsToTry) {
          try {
            await mappingList.items.add(p);
            mappingSuccess = true;
            console.log("Successfully added assignment record to Mapping List");
            break;
          } catch (err) {
            mappingLastErr = err;
          }
        }

        if (!mappingSuccess) {
          console.warn("Failed to write to Mapping List after trying all payloads. Continuing.", mappingLastErr);
        }
      } catch (err) {
        console.warn("Failed to execute Mapping List update logic. Continuing.", err);
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

  public static async syncExistingAssignmentsToMappingList(adminName: string): Promise<void> {
    try {
      console.log("Starting Mapping List sync...");
      const list = await InventoryService.getInventoryList();
      const items = await InventoryService.getItems();
      
      // Filter for assigned items matching the 5 active employees
      const assignedItems = items.filter(item => {
        const statusLower = (item.status || '').toLowerCase();
        const isAssigned = statusLower === 'assigned' || statusLower.indexOf('assigned to') >= 0;
        if (!isAssigned) return false;

        const employeeName = item.assignedTo || "";
        return EMPLOYEES.some(e => e.name.toLowerCase() === employeeName.toLowerCase());
      });

      if (assignedItems.length === 0) {
        console.log("No assigned assets found for the 5 active employees.");
        return;
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

      const employeFieldName = findMappingField("employe", "Employe");
      const employeeIdFieldName = findMappingField("employeid", "") || findMappingField("employeeid", "EmployeeID");
      const assetNameFieldName = findMappingField("assetname", "AssetName");
      const serialNumberFieldName = findMappingField("serialnumber", "SerialNumber");
      const priorityFieldName = findMappingField("priority", "Priority");
      const requestedDateFieldName = findMappingField("requesteddate", "RequestedDate");
      const reasonFieldName = findMappingField("reasonforrequest", "ReasonforRequest");
      const assignedDateFieldName = findMappingField("assigneddate", "AssignedDate");

      const employeFieldObj = mappingFields.find((f: any) => f.InternalName === employeFieldName);
      const isEmployePerson = employeFieldObj && (employeFieldObj.TypeAsString === "User" || employeFieldObj.TypeAsString === "UserMulti");

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

        const assetAssignedTo = asset.assignedTo || "";
        console.log(`Syncing missing assigned asset to Mapping List: ${asset.assetName || asset.title} (${asset.serialNumber}) for ${assetAssignedTo}`);

        const employee = EMPLOYEES.find(e => e.name.toLowerCase() === assetAssignedTo.toLowerCase());
        const employeeName = employee ? employee.name : assetAssignedTo;
        const employeeId = employee ? employee.id : "";
        const employeeEmail = employee ? employee.email : "";

        // Resolve Person ID if the field is a Person field
        let assignedToId: number | null = null;
        if (isEmployePerson && employeeEmail) {
          try {
            const sp = getSP();
            const user: any = await sp.web.ensureUser(employeeEmail);
            assignedToId = user.data ? user.data.Id : user.Id;
          } catch (e) {
            console.warn(`Could not resolve user ${employeeEmail} in SharePoint during sync.`, e);
          }
        }

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

        const dynamicMappingPayload: any = {};
        if (employeFieldName === "Title") {
          dynamicMappingPayload["Title"] = employeeName;
        } else {
          dynamicMappingPayload["Title"] = `Assignment of ${assetName}`;
          if (isEmployePerson && assignedToId !== null) {
            dynamicMappingPayload[`${employeFieldName}Id`] = assignedToId;
          } else {
            dynamicMappingPayload[employeFieldName] = employeeName;
          }
        }

        dynamicMappingPayload[employeeIdFieldName] = employeeId || "";
        dynamicMappingPayload[assetNameFieldName] = assetName;
        dynamicMappingPayload[serialNumberFieldName] = serialNumber;
        dynamicMappingPayload[priorityFieldName] = priority;
        dynamicMappingPayload[requestedDateFieldName] = finalRequestedDate;
        dynamicMappingPayload[reasonFieldName] = reason;
        dynamicMappingPayload[assignedDateFieldName] = finalAssignedDate;

        const mappingPayloadsToTry: any[] = [];
        mappingPayloadsToTry.push(dynamicMappingPayload);

        if (assignedToId !== null) {
          mappingPayloadsToTry.push({
            Title: employeeName,
            EmployeId: assignedToId,
            EmployeeID: employeeId || "",
            AssetName: assetName,
            SerialNumber: serialNumber,
            Priority: priority,
            RequestedDate: finalRequestedDate,
            ReasonforRequest: reason,
            AssignedDate: finalAssignedDate
          });
          mappingPayloadsToTry.push({
            Title: `Assignment of ${assetName}`,
            EmployeId: assignedToId,
            EmployeeID: employeeId || "",
            AssetName: assetName,
            SerialNumber: serialNumber,
            Priority: priority,
            RequestedDate: finalRequestedDate,
            ReasonforRequest: reason,
            AssignedDate: finalAssignedDate
          });
          mappingPayloadsToTry.push({
            Title: `Assignment of ${assetName}`,
            EmployeId: assignedToId,
            Employee_x0020_ID: employeeId || "",
            Asset_x0020_Name: assetName,
            Serial_x0020_Number: serialNumber,
            Priority: priority,
            Requested_x0020_Date: finalRequestedDate,
            Reason_x0020_for_x0020_Request: reason,
            AssignedDate: finalAssignedDate
          });
        }

        mappingPayloadsToTry.push({
          Title: employeeName,
          Employe: employeeName,
          EmployeeID: employeeId || "",
          AssetName: assetName,
          SerialNumber: serialNumber,
          Priority: priority,
          RequestedDate: finalRequestedDate,
          ReasonforRequest: reason,
          AssignedDate: finalAssignedDate
        });
        mappingPayloadsToTry.push({
          Title: `Assignment of ${assetName}`,
          Employe: employeeName,
          EmployeeID: employeeId || "",
          AssetName: assetName,
          SerialNumber: serialNumber,
          Priority: priority,
          RequestedDate: finalRequestedDate,
          ReasonforRequest: reason,
          AssignedDate: finalAssignedDate
        });
        mappingPayloadsToTry.push({
          Title: `Assignment of ${assetName}`,
          Employee: employeeName,
          EmployeeID: employeeId || "",
          AssetName: assetName,
          SerialNumber: serialNumber,
          Priority: priority,
          RequestedDate: finalRequestedDate,
          ReasonforRequest: reason,
          AssignedDate: finalAssignedDate
        });

        let mappingSuccess = false;
        let mappingLastErr: any;
        for (const p of mappingPayloadsToTry) {
          try {
            await mappingList.items.add(p);
            mappingSuccess = true;
            console.log(`Successfully synced record to Mapping List for ${employeeName}`);
            break;
          } catch (err) {
            mappingLastErr = err;
          }
        }

        if (!mappingSuccess) {
          console.warn("Failed to write synced record to Mapping List. Continuing.", mappingLastErr);
        }
      }
    } catch (e) {
      console.error("Failed to sync existing assignments to Mapping List:", e);
    }
  }

  public static async getAuditLogs(): Promise<IEventLog[]> {
    let logs: IEventLog[] = [];
    try {
      // 1. Fetch from InventoryList
      try {
        const list = await InventoryService.getInventoryList();
        const inventoryItems = await list.items.select("ID", "Title", "AssetName", "Created", "Modified", "Author/Title", "Editor/Title").expand("Author", "Editor")();
        
        inventoryItems.forEach((item: any) => {
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
        const reqList = await InventoryService.getRequestList();
        const requestItems = await this._fetchItemsWithExpandedUsers(reqList);
        
        requestItems.forEach((item: any) => {
          const keys = Object.keys(item);
          const findKey = (searchStr: string) => {
            // Prefer matching keys that do not end in "id" (like "Requester" over "RequesterId")
            const nonIdMatch = keys.find(k => {
              const kl = k.toLowerCase().replace(/_x0020_/g, '');
              return kl.indexOf(searchStr) >= 0 && !kl.endsWith("id");
            });
            if (nonIdMatch) return nonIdMatch;
            return keys.find(k => k.toLowerCase().replace(/_x0020_/g, '').indexOf(searchStr) >= 0);
          };

           const employeeKey = findKey("requester") || findKey("employee") || "Employee";
           const selectAssetKey = findKey("assettype") || findKey("selectasset") || findKey("type") || "SelectAsset";

           const reqAssetName = item[selectAssetKey] || item.Title || "Unknown Asset";
           const rawEmp = item[employeeKey] || item.Employee || item.Author;
           const reqUser = (() => {
             if (!rawEmp) return item.Title || "System";
             if (typeof rawEmp === 'string') return rawEmp;
             if (Array.isArray(rawEmp)) return rawEmp.map((a: any) => a.Title || a.Name || "").join(', ');
             if (typeof rawEmp === 'object') return rawEmp.Title || rawEmp.Name || JSON.stringify(rawEmp);
             return rawEmp.toString();
           })();

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
    await this._ensureRequestWorkflowFields();

    const normalizedRequestKey = this._normalizeRequestKey(requestLookupId);
    if (!normalizedRequestKey) {
      throw new Error("Request ID is required.");
    }

    const reqList = await InventoryService.getRequestList();
    let requestItems: any[] = [];
    try {
      requestItems = await reqList
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
}
