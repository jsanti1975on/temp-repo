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
