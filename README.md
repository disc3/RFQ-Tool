# RFQ Tool - Project Documentation

## Key Features

### 1. SAP Automation & Integration

The tool automates several SAP transactions via SAP GUI Scripting:

- **BOM Creation (`CreateBOM.bas`)**: Automates `CS01` for creating Bill of Materials.
- **Routing Creation (`CreateRouting.bas`)**: Automates `CA01` for creating Routings.
- **Component Allocation (`AddComponentAllocation.bas`)**: Automates `CA02` to allocate components to operations.
- **Stock Check (`SAP_CO09_Exporter.bas`)**: Automates `CO09` to check provisional free stock for components.
- **Template Export**: Generates upload-ready Excel sheets for BOMs and Routings (`SAP_BOM_Uploader_copy.bas`, `SAP_Routing_Uploader_copy.bas`).

### 2. Routine & Variant Management

### Common Issues

- **"Server Name Not Found"**: Check network connection and Power Automate URL.
- **SAP Scripting Errors**: Ensure scripting is enabled in SAP GUI settings and the correct transaction is open if required.
an unrecognized format, which may limit the completeness of this manual.
