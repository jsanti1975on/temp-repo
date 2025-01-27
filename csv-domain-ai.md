# FLOWCHART POINT: Promotion Process Server-1 ANN01

# Primary Domain Controller: ANN01
```PowerShell
"Install ADDS and DNS roles. Promote to Domain Controller for MACHINE.AI. Set DSRM password and restart."
```

# Create OUs
```PowerShell
New-ADOrganizationalUnit -Name "MODELS" -Description "AI Models" -DisplayName "MODELS" -ProtectedFromAccidentalDeletion $True -Path "DC=machine,DC=ai"

New-ADOrganizationalUnit -Name "TERMS" -Description "AI Terms" -DisplayName "TERMS" -ProtectedFromAccidentalDeletion $True -Path "DC=machine,DC=ai"

New-ADOrganizationalUnit -Name "TYPES" -Description "AI Types" -DisplayName "TYPES" -ProtectedFromAccidentalDeletion $True -Path "DC=machine,DC=ai"

```


# Create Groups
```PowerShell
New-ADGroup -Name "G_models" -SamAccountName "G_models" -GroupCategory Security -GroupScope Global -DisplayName "G_models" -Path "OU=MODELS,DC=machine,DC=ai" -Description "Group for AI Models"

New-ADGroup -Name "G_terms" -SamAccountName "G_terms" -GroupCategory Security -GroupScope Global -DisplayName "G_terms" -Path "OU=TERMS,DC=machine,DC=ai" -Description "Group for AI Terms"

New-ADGroup -Name "G_types" -SamAccountName "G_types" -GroupCategory Security -GroupScope Global -DisplayName "G_types" -Path "OU=TYPES,DC=machine,DC=ai" -Description "Group for AI Types"

```


# Import Users (CSV):
```csv
File Name: csvusers.csv
File Contents:
dn,SamAccountName,userPrincipalName,objectClass
"cn=ChatGPT,ou=MODELS,dc=machine,dc=ai",ChatGPT,ChatGPT@machine.ai,user
"cn=OpenAI,ou=MODELS,dc=machine,dc=ai",OpenAI,OpenAI@machine.ai,user
"cn=Bing,ou=MODELS,dc=machine,dc=ai",Bing,Bing@machine.ai,user
"cn=Claude,ou=MODELS,dc=machine,dc=ai",Claude,Claude@machine.ai,user
"cn=Whisper,ou=TERMS,dc=machine,dc=ai",Whisper,Whisper@machine.ai,user
"cn=Hallucination,ou=TERMS,dc=machine,dc=ai",Hallucination,Hallucination@machine.ai,user
"cn=Turing,ou=TERMS,dc=machine,dc=ai",Turing,Turing@machine.ai,user
"cn=Prompt,ou=TERMS,dc=machine,dc=ai",Prompt,Prompt@machine.ai,user
"cn=Generative,ou=TYPES,dc=machine,dc=ai",Generative,Generative@machine.ai,user
"cn=Conversation,ou=TYPES,dc=machine,dc=ai",Conversation,Conversation@machine.ai,user
"cn=Deep,ou=TYPES,dc=machine,dc=ai",Deep,Deep@machine.ai,user
"cn=Limited,ou=TYPES,dc=machine,dc=ai",Limited,Limited@machine.ai,user
```
______________________________________> Run:
csvde -i -f csvusers.csv

## New Project

```
G_Cattleya
Hazel Boyd, hboyd
Ken Dream, kdream
Smiley Sunset, ssunset
----
G_Vanda
Antonio Romani, aromani
Mimi Palmer, mpalmer
Pedro Bonetti, pbonetti
---
G_Catasetum
Chuck Taylor, ctaylor
Fred Clark, fclark
Jean Monnier, jmonnier
```

## Clean template 

```csv
dn,SamAccountName,userPrincipalName,objectClass
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
"cn=<<Name>>,ou=<<OU>>,dc=domain,dc=name",<<Name>>,<<Name>>@<<Domain>>,user
```

## Template filled

```csv
dn,SamAccountName,userPrincipalName,objectClass
"cn=G_Cattleya,ou=Cattleya,dc=orkid-west,dc=arpa",G_Cattleya,G_Cattleya@orkid-west.arpa,user
"cn=Hazel Boyd,ou=Cattleya,dc=orkid-west,dc=arpa",hboyd,hboyd@orkid-west.arpa,user
"cn=Ken Dream,ou=Cattleya,dc=orkid-west,dc=arpa",kdream,kdream@orkid-west.arpa,user
"cn=Smiley Sunset,ou=Cattleya,dc=orkid-west,dc=arpa",ssunset,ssunset@orkid-west.arpa,user
"cn=G_Vanda,ou=Vanda,dc=orkid-west,dc=arpa",G_Vanda,G_Vanda@orkid-west.arpa,user
"cn=Antonio Romani,ou=Vanda,dc=orkid-west,dc=arpa",aromani,aromani@orkid-west.arpa,user
"cn=Mimi Palmer,ou=Vanda,dc=orkid-west,dc=arpa",mpalmer,mpalmer@orkid-west.arpa,user
"cn=Pedro Bonetti,ou=Vanda,dc=orkid-west,dc=arpa",pbonetti,pbonetti@orkid-west.arpa,user
"cn=G_Catasetum,ou=Catasetum,dc=orkid-west,dc=arpa",G_Catasetum,G_Catasetum@orkid-west.arpa,user
"cn=Chuck Taylor,ou=Catasetum,dc=orkid-west,dc=arpa",ctaylor,ctaylor@orkid-west.arpa,user
"cn=Fred Clark,ou=Catasetum,dc=orkid-west,dc=arpa",fclark,fclark@orkid-west.arpa,user
"cn=Jean Monnier,ou=Catasetum,dc=orkid-west,dc=arpa",jmonnier,jmonnier@orkid-west.arpa,user
```

```TypeScript
function main(workbook: ExcelScript.Workbook) {
  // Define the active worksheet
  const sheet = workbook.getActiveWorksheet();
  
  // Define the domain name (this can be modified as needed)
  const domainName = "orkid-west.arpa";

  // Mapping for Global Security Groups
  const groupMappings = {
    "G_Cattleya": ["Hazel", "Ken", "Smiley"],
    "G_Vanda": ["Antonio", "Mimi", "Pedro"],
    "G_Catasetum": ["Chuck", "Fred", "Jean"]
  };

  // Define the range where user data is filled (assumed starting from row 2)
  const firstNameRange = sheet.getRange("A2:A100");
  const lastNameRange = sheet.getRange("B2:B100");
  const rows = firstNameRange.getValues();  // Get the data from the First Name column

  // Prepare the CSV header
  let csvContent = "First Name,Last Name,Domain Name,Global Security Group,SamAccountName,userPrincipalName\n";

  // Loop through the user data and process each row
  rows.forEach((row, index) => {
    const firstName = row[0];  // First Name from column A
    const lastName = lastNameRange.getCell(index, 0).getValue();  // Last Name from column B
    
    // Check if First Name and Last Name are filled in
    if (firstName && lastName) {
      // Determine the Global Security Group based on the First Name (this is just an example logic)
      let securityGroup = "";
      for (const [group, members] of Object.entries(groupMappings)) {
        if (members.includes(firstName)) {
          securityGroup = group;
          break;
        }
      }

      // Generate the SamAccountName (e.g., first initial + last name)
      const samAccountName = firstName.charAt(0).toLowerCase() + lastName.toLowerCase();
      
      // Generate the User Principal Name (email)
      const userPrincipalName = `${samAccountName}@${domainName}`;

      // Add the row to the CSV content
      csvContent += `${firstName},${lastName},${domainName},${securityGroup},${samAccountName},${userPrincipalName}\n`;

      // Optionally, you can fill in the Domain Name and Security Group into the worksheet
      sheet.getRange(`C${index + 2}`).setValue(domainName);
      sheet.getRange(`D${index + 2}`).setValue(securityGroup);
      sheet.getRange(`E${index + 2}`).setValue(samAccountName);
      sheet.getRange(`F${index + 2}`).setValue(userPrincipalName);
    }
  });

  // Log the CSV content to the console (you can download it or handle it further as needed)
  console.log(csvContent);

  // Optional: Notify the user that the CSV has been generated
  workbook.getApplication().showNotification("CSV file generated. Check the console output.");
}
```
