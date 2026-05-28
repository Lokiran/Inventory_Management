import os

file_path = r"c:\Users\00325102\OneDrive - Nexer AB\Desktop\INVENTORY\src\webparts\employeeManagement\components\AssetRequest\AssetRequestModule.tsx"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

target = """              <TextField
                label="Employee Name"
                value={formData.employeeName}
                disabled
              />"""

replacement = """              <TextField
                label="Employee Name"
                value={formData.employeeName}
                onChange={(ev, newValue) => handleInputChange('employeeName', newValue)}
              />"""

# Handling \r\n vs \n
target = target.replace('\r\n', '\n')
replacement = replacement.replace('\r\n', '\n')
content_norm = content.replace('\r\n', '\n')

content_new = content_norm.replace(target, replacement)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content_new)
