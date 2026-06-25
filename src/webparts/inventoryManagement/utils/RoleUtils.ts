export type UserRole = 'Admin' | 'Inventory Manager' | 'Inventory Employee';

export const RoleUtils = {
  getRoleLevel: (role: UserRole): number => {
    switch (role) {
      case 'Admin':
        return 2;
      case 'Inventory Manager':
        return 1;
      case 'Inventory Employee':
        return 0;
      default:
        return 0;
    }
  },

  canAddAssets: (role: UserRole): boolean => {
    return role === 'Admin' || role === 'Inventory Manager';
  },

  canApproveRequests: (role: UserRole): boolean => {
    return role === 'Inventory Manager' || role === 'Admin';
  },

  canViewAuditLogs: (role: UserRole): boolean => {
    return role === 'Admin' || role === 'Inventory Manager';
  },

  canAssignAssets: (role: UserRole): boolean => {
    return role === 'Admin' || role === 'Inventory Manager';
  },

  canManageUsers: (role: UserRole): boolean => {
    return role === 'Admin';
  },

  canAccessReports: (role: UserRole): boolean => {
    return role === 'Admin';
  },

  canAccessConfig: (role: UserRole): boolean => {
    return role === 'Admin';
  },

  isAdmin: (role: UserRole): boolean => {
    return role === 'Admin';
  },

  isManager: (role: UserRole): boolean => {
    return role === 'Inventory Manager';
  },

  isEmployee: (role: UserRole): boolean => {
    return role === 'Inventory Employee';
  },

  hasPermission: (role: UserRole, minRequiredRole: UserRole): boolean => {
    return RoleUtils.getRoleLevel(role) >= RoleUtils.getRoleLevel(minRequiredRole);
  },

  getRoleDisplayName: (role: UserRole): string => {
    return role;
  }
};
