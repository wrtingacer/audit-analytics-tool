# AuditXcel AI - MVP

This repo contains the Excel add-in for AuditXcel AI.

## Files:
- MainModule.bas: Core functions like cleaning, summaries, etc.
- RibbonCallbacks.cls: Handles ribbon button clicks.
- RibbonXML.bas: Defines the custom ribbon.
- AuditXcelAI.xlam: The compiled add-in file—download and install in Excel.

## How to Use:
1. Download AuditXcelAI.xlam.
2. In Excel: File > Options > Add-ins > Manage: Excel Add-ins > Go > Browse to file > OK.
3. Test the "AuditXcel AI" tab.

## To Build from Source:
- Create a new .xlam in Excel.
- VBA Editor: Import .bas and .cls files.
- Use Custom UI Editor to add ribbon XML.
