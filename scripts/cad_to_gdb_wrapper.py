"""
===============================================================================
NMDOT CAD → GIS Conversion Tool (V1)
===============================================================================

Author: Narendra Lingutla
Organization: NMDOT / Srirama LLC
Version: 1.0
ArcGIS Pro: 3.x+
License: ArcGIS Pro Basic or higher

-------------------------------------------------------------------------------
PURPOSE
-------------------------------------------------------------------------------
This tool converts CAD datasets (DWG, DGN, DXF) into a GIS-ready File
Geodatabase using ArcGIS Pro's CAD To Geodatabase conversion engine.

The tool is designed as a product-ready, no-code workflow for DOT users and:
- Copies a standardized template geodatabase
- Creates a per-run staging Feature Dataset (CAD_RAW)
- Defines the spatial reference (does NOT project CAD data)
- Optionally adds results to the active ArcGIS Pro map

-------------------------------------------------------------------------------
IMPORTANT GIS RULE (READ THIS)
-------------------------------------------------------------------------------
CAD data is assumed to already be in a coordinate system.
This tool DEFINES the coordinate system — it does NOT reproject CAD.

If the wrong .prj is supplied, the data will be misaligned.

-------------------------------------------------------------------------------
WORKFLOW OVERVIEW
-------------------------------------------------------------------------------
1. Validate input CAD folder and projection file
2. Copy template.gdb → <OutputFolder>/<Project>_Converted_<timestamp>.gdb
3. Create staging Feature Dataset (CAD_RAW_<Project>_<timestamp>)
4. Run CAD To Geodatabase into the staging dataset
5. (Optional) Auto-add results to the active ArcGIS Pro map

-------------------------------------------------------------------------------
DESIGN NOTES
-------------------------------------------------------------------------------
• A new output geodatabase is created per run to avoid name collisions
• Staging datasets are unique per run (future option to overwrite exists)
• Intended as V1 foundation for V2 schema mapping & QC workflows

-------------------------------------------------------------------------------
LIMITATIONS
-------------------------------------------------------------------------------
• CAD geometry quality is not modified
• No attribute remapping in V1
• Annotation is imported as-is

===============================================================================
"""


import arcpy
import os
import shutil
from datetime import datetime

arcpy.env.overwriteOutput = True

def _clean_name(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return "CAD_Project"

    # ArcGIS-safe: letters, numbers, underscore ONLY
    cleaned = "".join(c if c.isalnum() else "_" for c in s)

    # Cannot start with a number
    if cleaned[0].isdigit():
        cleaned = f"X_{cleaned}"

    # Trim excessive underscores
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")

    return cleaned.strip("_")[:64]

def _find_cad_files(folder: str, recursive: bool) -> list[str]:
    exts = (".dwg", ".dgn", ".dxf")
    cad_files = []

    if recursive:
        for root, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(exts):
                    cad_files.append(os.path.join(root, f))
    else:
        for f in os.listdir(folder):
            if f.lower().endswith(exts):
                cad_files.append(os.path.join(folder, f))

    return sorted(cad_files)


def _unique_gdb_path(output_folder: str, base_name: str) -> str:
    """
    Creates a UNIQUE output gdb per run to avoid name collisions like Point/Polyline/Polygon.
    Example:
      <OutputFolder>\<base>_Converted_20251215_190010.gdb
    """
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    gdb_name = f"{base_name}_Converted_{stamp}.gdb"
    return os.path.join(output_folder, gdb_name)

def _copy_template_gdb(template_gdb: str, target_gdb: str):
    # File GDB is a folder; copytree is reliable if target doesn't exist.
    if os.path.exists(target_gdb):
        shutil.rmtree(target_gdb)  
    shutil.copytree(template_gdb, target_gdb)

def _add_dataset_to_map(dataset_path: str):
    try:
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        m = aprx.activeMap
        if m is None:
            arcpy.AddWarning(
                "No active map found. To auto-add results, open or create a map "
                "in the current ArcGIS Pro project before running the tool."
                )
            return
        m.addDataFromPath(dataset_path)
        arcpy.AddMessage("Added output dataset to the active map.")
    except Exception as e:
        arcpy.AddMessage(f"Auto-add to map skipped: {e}")

def _warn_if_extent_suspicious(dataset_path: str):
    try:
        desc = arcpy.Describe(dataset_path)
        ext = desc.extent
        width = abs(ext.XMax - ext.XMin)
        height = abs(ext.YMax - ext.YMin)

        # crude but effective checks
        if width > 5_000_000 or height > 5_000_000:
            arcpy.AddWarning(
                "Output extent is extremely large. "
                "This usually indicates an incorrect coordinate system definition."
            )
    except Exception:
        pass

def main():
    try:
        # --- Script Tool Parameters ---
        # 0 Input CAD Folder (DEFolder)
        # 1 Output Folder (DEFolder)
        # 2 Projection File (.prj) (DEFile)
        # 3 Project Name (optional string)  [can be empty]
        # 4 Auto-add to map (optional boolean) [default True in tool]
        cad_folder = arcpy.GetParameterAsText(0)
        output_folder = arcpy.GetParameterAsText(1)
        prj_file = arcpy.GetParameterAsText(2)
        project_name = arcpy.GetParameterAsText(3)
        auto_add = arcpy.GetParameter(4)  # boolean
        search_subfolders = arcpy.GetParameter(5)  # boolean

        # # NOTE:
        # # CADToGeodatabase creates feature classes named Point/Polyline/Polygon.
        # # These must be unique across the entire GDB.
        # # Therefore:
        # # - Fixed staging dataset ONLY works if output GDB is overwritten each run
        # # - Per-run staging dataset REQUIRES a new output GDB per run (current default)
        # USE_FIXED_STAGING_DATASET = False  # FUTURE OPTION
        # if USE_FIXED_STAGING_DATASET:
        #     staging_dataset_name = "CAD_RAW"
        # else:
        #     stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        #     staging_dataset_name = f"CAD_RAW_{project_name}_{stamp}"


        if not cad_folder or not os.path.isdir(cad_folder):
            arcpy.AddError("Input CAD Folder is invalid or missing.")
            return

        # Default project name = CAD folder name
        if not project_name:
            project_name = os.path.basename(os.path.normpath(cad_folder))
        project_name = _clean_name(project_name)

        cad_files = _find_cad_files(cad_folder, search_subfolders)
        if not cad_files:
            msg = "No CAD files (.dwg/.dgn/.dxf) found in the input folder."
            if not search_subfolders:
                msg += " Tip: If your CAD files are inside subfolders, enable 'Search subfolders'."
            arcpy.AddError(msg)
            return
        arcpy.AddMessage(f"Found {len(cad_files)} CAD files (recursive={search_subfolders})")

        if not output_folder:
            arcpy.AddError("Output Folder is required.")
            return
        os.makedirs(output_folder, exist_ok=True)

        if not prj_file or not os.path.isfile(prj_file):
            arcpy.AddWarning("No projection file provided. CAD data will be imported WITHOUT a defined coordinate system.")
            sr = None
        else:
            sr = arcpy.SpatialReference(prj_file)

        if sr:
            arcpy.AddMessage("--------------------------------------------------")
            arcpy.AddMessage("CAD Coordinate System Definition")
            arcpy.AddMessage("--------------------------------------------------")
            arcpy.AddMessage(f"Name: {sr.name}")
            arcpy.AddMessage(f"Factory Code: {sr.factoryCode}")
            arcpy.AddMessage(f"Linear Unit: {sr.linearUnitName}")
            arcpy.AddMessage(f"Datum: {sr.GCS.name if sr.GCS else 'Unknown'}")
        else:
            arcpy.AddMessage("No spatial reference will be defined on import.") 

        script_dir = os.path.dirname(os.path.abspath(__file__))
        template_gdb = os.path.join(script_dir, "template.gdb")
        if not os.path.isdir(template_gdb):
            arcpy.AddError(f"template.gdb not found next to script: {template_gdb}")
            return

        # # CAD list
        # cad_files = [
        #     os.path.join(cad_folder, f)
        #     for f in os.listdir(cad_folder)
        #     if f.lower().endswith((".dwg", ".dgn", ".dxf"))
        # ]
        if not cad_files:
            arcpy.AddError("No CAD files (.dwg/.dgn/.dxf) found in the input folder.")
            return
        
        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage("Starting CAD → Geodatabase conversion")
        arcpy.AddMessage("--------------------------------------------------")

        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage("STEP 1: Validating inputs")
        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage(f"Found {len(cad_files)} CAD files")
        arcpy.AddMessage(f"Project Name: {project_name}")

        # Unique output GDB per run (required for staging-per-run design)
        out_gdb = _unique_gdb_path(output_folder, project_name)

        # Copy template → output GDB
        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage("STEP 2: Creating output Geodatabase from template")
        arcpy.AddMessage("--------------------------------------------------")

        arcpy.AddMessage(f"Creating output GDB from template: {out_gdb}")
        _copy_template_gdb(template_gdb, out_gdb)

        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage("STEP 3: Creating staging Feature Dataset")
        arcpy.AddMessage("--------------------------------------------------")

        # Unique staging dataset per run
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        staging_dataset_name = f"CAD_RAW_{project_name}_{stamp}"
        staging_dataset_name = _clean_name(staging_dataset_name)

        arcpy.AddMessage(f"Creating staging Feature Dataset: {staging_dataset_name}")
        if sr:
            arcpy.management.CreateFeatureDataset(out_gdb, staging_dataset_name, sr)
        else:
            arcpy.management.CreateFeatureDataset(out_gdb, staging_dataset_name)

        # Run CADToGeodatabase into the output GDB, writing to the staging dataset
        cad_inputs = ";".join(cad_files)
        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage("STEP 4: CAD to Geodatabase Conversion")
        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage("Running CADToGeodatabase...")
        arcpy.conversion.CADToGeodatabase(
            input_cad_datasets=cad_inputs,
            out_gdb_path=out_gdb,
            out_dataset_name=staging_dataset_name,
            reference_scale=1000,
            spatial_reference=sr
        )

        staging_path = os.path.join(out_gdb, staging_dataset_name)
        _warn_if_extent_suspicious(staging_path)
        arcpy.AddMessage("CAD To Geodatabase completed successfully.")
        arcpy.AddMessage(f"Output GDB: {out_gdb}")
        arcpy.AddMessage(f"Staging Dataset: {staging_path}")

        arcpy.AddMessage("--------------------------------------------------")
        arcpy.AddMessage("STEP 5: Map addition (optional)")
        arcpy.AddMessage("--------------------------------------------------")

        if auto_add:
            _add_dataset_to_map(staging_path)
    
    except arcpy.ExecuteError:
        arcpy.AddError("ArcGIS geoprocessing error occurred.")
        arcpy.AddError(arcpy.GetMessages(2))
        raise
    except Exception as ex:
        arcpy.AddError("Unexpected error occurred.")
        arcpy.AddError(str(ex))
        raise

    arcpy.AddWarning(
        "Reminder: CAD coordinates are NOT transformed. "
        "If output is misaligned, verify the CAD drawing's coordinate system "
        "in Civil 3D or OpenRoads before rerunning."
    )

    # Optional derived outputs for ModelBuilder chaining
    # (Create 2 derived output parameters in the script tool if you want)
    arcpy.SetParameterAsText(6, out_gdb)         # Derived Output GDB path
    arcpy.SetParameterAsText(7, staging_path)    # Derived Output Dataset path

if __name__ == "__main__":
    main()
