# NMDOT CAD → GIS Conversion Tool (V1)

A production-ready, ModelBuilder-driven CAD to GIS conversion tool for ArcGIS Pro.

## Overview

This tool converts roadway design CAD files (DWG, DGN, DXF) into GIS-ready
File Geodatabases using ArcGIS Pro’s native **CAD To Geodatabase** engine.

It is designed for DOT GIS analysts, planners, and engineers and requires
**no Python knowledge** to operate.

## Key Features

- No-code execution via ArcGIS Pro Script Tool
- Per-run output File Geodatabase
- Per-run staging Feature Dataset (`CAD_RAW`)
- Correct coordinate system definition (no reprojection)
- Recursive CAD file discovery
- Optional auto-add to map

## Important GIS Rule

> CAD data is assumed to already be in a coordinate system.
> This tool DEFINES the coordinate system — it does NOT reproject data.

Using the wrong `.prj` file will result in misaligned output.

## Requirements

- ArcGIS Pro 3.x+
- ArcGIS Pro Basic or higher license
- Windows OS

## Folder Structure

NMDOT_CAD2GIS_PRODUCT/
│
├── scripts/
│ └── cad_to_gdb_wrapper.py
├── template.gdb/
├── toolbox/
│ └── NMDOT_CAD2GIS.tbx
├── docs/
│ └── NMDOT_CAD2GIS_UserGuide.docx
└── README.md

## Usage

1. Open ArcGIS Pro
2. Add the toolbox
3. Run `CADToGDBWrapper`
4. Provide input CAD folder, output folder, and projection file
5. Run

## Version

- **V1.0.0** — Foundation release

## License

Internal use for NMDOT and SRIRAMA, LLC.
