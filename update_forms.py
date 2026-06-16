import os

# Update AssetRequestModule.tsx
asset_path = r"c:\Users\00325102\OneDrive - Nexer AB\Desktop\INVENTORY\src\webparts\employeeManagement\components\AssetRequest\AssetRequestModule.tsx"

with open(asset_path, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Update interface
content = content.replace("  employeeId: string;", "  employeeEmail: string;")

# 2. Update initial state and resets (there are two places: useState and handleCancel / reset)
content = content.replace("employeeId: '',", "employeeEmail: '',")

# 3. Update the validation check
content = content.replace("if (!formData.employeeId", "if (!formData.employeeEmail")

# 4. Update the payload to use formData.employeeEmail instead of props.userEmail
old_payload = """      const payload = {
        employeeEmail: props.userEmail,
        assetType: formData.assetType,"""
new_payload = """      const payload = {
        employeeEmail: formData.employeeEmail,
        assetType: formData.assetType,"""
content = content.replace(old_payload, new_payload)

# 5. Update the UI TextField
old_textfield = """              <TextField
                label="Employee ID *"
                placeholder="Enter your employee ID"
                value={formData.employeeId}
                onChange={(ev, newValue) => handleInputChange('employeeId', newValue)}
                required
              />"""
new_textfield = """              <TextField
                label="Employee Email *"
                placeholder="e.g., john.doe@company.com"
                value={formData.employeeEmail}
                onChange={(ev, newValue) => handleInputChange('employeeEmail', newValue)}
                required
              />"""
content = content.replace(old_textfield, new_textfield)

# Handling CRLF issues just in case
content_norm = content.replace('\r\n', '\n')
old_payload_norm = old_payload.replace('\r\n', '\n')
new_payload_norm = new_payload.replace('\r\n', '\n')
old_textfield_norm = old_textfield.replace('\r\n', '\n')
new_textfield_norm = new_textfield.replace('\r\n', '\n')

content_norm = content_norm.replace(old_payload_norm, new_payload_norm)
content_norm = content_norm.replace(old_textfield_norm, new_textfield_norm)

with open(asset_path, 'w', encoding='utf-8') as f:
    f.write(content_norm)

# Update IncidentRequestModule.tsx
incident_path = r"c:\Users\00325102\OneDrive - Nexer AB\Desktop\INVENTORY\src\webparts\employeeManagement\components\IncidentRequest\IncidentRequestModule.tsx"

with open(incident_path, 'r', encoding='utf-8') as f:
    inc_content = f.read()

# 1. Add employeeEmail to interface
inc_content = inc_content.replace("  assetId: string;", "  employeeEmail: string;\n  assetId: string;")

# 2. Add employeeEmail to initial state and resets
inc_content = inc_content.replace("assetId: '',", "employeeEmail: '',\n    assetId: '',")

# 3. Update the validation check
inc_content = inc_content.replace("if (!formData.assetId", "if (!formData.employeeEmail || !formData.assetId")

# 4. Update the payload to use formData.employeeEmail instead of props.userEmail
old_inc_payload = """      const payload = {
        ...formData,
        employeeEmail: props.userEmail
      };"""
new_inc_payload = """      const payload = {
        ...formData,
        employeeEmail: formData.employeeEmail
      };"""
inc_content = inc_content.replace(old_inc_payload, new_inc_payload)

# 5. Add UI TextField for Employee Email
old_inc_asset_info = """              <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                Asset Information
              </Text>"""
new_inc_asset_info = """              <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c' }}>
                Employee Information
              </Text>
              <TextField
                label="Employee Email *"
                placeholder="e.g., john.doe@company.com"
                value={formData.employeeEmail}
                onChange={(ev, newValue) => handleInputChange('employeeEmail', newValue)}
                required
              />
              <Text variant="large" style={{ fontWeight: 600, color: '#e74c3c', marginTop: '15px' }}>
                Asset Information
              </Text>"""

inc_content_norm = inc_content.replace('\r\n', '\n')
old_inc_payload_norm = old_inc_payload.replace('\r\n', '\n')
new_inc_payload_norm = new_inc_payload.replace('\r\n', '\n')
old_inc_asset_info_norm = old_inc_asset_info.replace('\r\n', '\n')
new_inc_asset_info_norm = new_inc_asset_info.replace('\r\n', '\n')

inc_content_norm = inc_content_norm.replace(old_inc_payload_norm, new_inc_payload_norm)
inc_content_norm = inc_content_norm.replace(old_inc_asset_info_norm, new_inc_asset_info_norm)

with open(incident_path, 'w', encoding='utf-8') as f:
    f.write(inc_content_norm)

print("Updates applied to AssetRequestModule and IncidentRequestModule")
