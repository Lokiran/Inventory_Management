export class InventoryService {
  private apiBaseUrl: string;

  constructor(apiBaseUrl: string) {
    this.apiBaseUrl = apiBaseUrl;
  }

  // Dashboard
  public async getDashboardStats(userEmail: string): Promise<any> {
    const response = await fetch(
      `${this.apiBaseUrl}/dashboard/stats?email=${encodeURIComponent(userEmail)}`,
      {
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.getToken()}`,
        },
      }
    );
    return this.handleResponse(response);
  }

  // Asset Requests
  public async createAssetRequest(requestData: any): Promise<any> {
    const response = await fetch(`${this.apiBaseUrl}/asset-requests`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${this.getToken()}`,
      },
      body: JSON.stringify(requestData),
    });
    return this.handleResponse(response);
  }

  public async getEmployeeAssetRequests(userEmail: string): Promise<any[]> {
    const response = await fetch(
      `${this.apiBaseUrl}/asset-requests/employee?email=${encodeURIComponent(userEmail)}`,
      {
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.getToken()}`,
        },
      }
    );
    return this.handleResponse(response);
  }

  public async cancelAssetRequest(requestId: string): Promise<any> {
    const response = await fetch(`${this.apiBaseUrl}/asset-requests/${requestId}/cancel`, {
      method: 'PUT',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${this.getToken()}`,
      },
    });
    return this.handleResponse(response);
  }

  // Incident Requests
  public async createIncidentRequest(incidentData: any, file?: File): Promise<any> {
    const formData = new FormData();
    if (incidentData.employeeEmail) {
      formData.append('employeeEmail', incidentData.employeeEmail);
    }
    formData.append('assetId', incidentData.assetId);
    formData.append('assetName', incidentData.assetName);
    formData.append('issueType', incidentData.issueType);
    formData.append('issueDescription', incidentData.issueDescription);
    formData.append('priority', incidentData.priority);
    formData.append('reportedDate', incidentData.reportedDate);

    if (file) {
      formData.append('attachment', file);
    }

    const response = await fetch(`${this.apiBaseUrl}/incidents`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${this.getToken()}`,
      },
      body: formData,
    });
    return this.handleResponse(response);
  }

  public async getEmployeeIncidentHistory(userEmail: string): Promise<any[]> {
    const response = await fetch(
      `${this.apiBaseUrl}/incidents/employee?email=${encodeURIComponent(userEmail)}`,
      {
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.getToken()}`,
        },
      }
    );
    return this.handleResponse(response);
  }

  // Assigned Assets
  public async getEmployeeAssignedAssets(userEmail: string): Promise<any[]> {
    const response = await fetch(
      `${this.apiBaseUrl}/assets/assigned?email=${encodeURIComponent(userEmail)}`,
      {
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${this.getToken()}`,
        },
      }
    );
    return this.handleResponse(response);
  }

  // Utility Methods
  private async handleResponse(response: Response): Promise<any> {
    if (!response.ok) {
      const error = await response.json().catch(() => ({ message: response.statusText }));
      throw new Error(error.message || `API Error: ${response.status}`);
    }
    return response.json();
  }

  private getToken(): string {
    // In a real application, retrieve the token from session storage or context
    // For now, return a placeholder
    return 'your-auth-token';
  }
}
