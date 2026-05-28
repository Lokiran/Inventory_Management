import os

file_path = r"c:\Users\00325102\OneDrive - Nexer AB\Desktop\INVENTORY\src\webparts\employeeManagement\components\IncidentRequest\IncidentRequestModule.tsx"

with open(file_path, 'r', encoding='utf-8') as f:
    content = f.read()

target = """      setIsSubmitting(true);
      const service = new InventoryService(props.apiBaseUrl);
      await service.createIncidentRequest(formData, selectedFile || undefined);"""

replacement = """      setIsSubmitting(true);
      const service = new InventoryService(props.apiBaseUrl);
      
      const payload = {
        ...formData,
        employeeEmail: props.userEmail
      };
      await service.createIncidentRequest(payload, selectedFile || undefined);"""

# Normalize
target = target.replace('\r\n', '\n')
replacement = replacement.replace('\r\n', '\n')
content_norm = content.replace('\r\n', '\n')

content_new = content_norm.replace(target, replacement)

with open(file_path, 'w', encoding='utf-8') as f:
    f.write(content_new)
