# 📦 BOM Processing Tool for Fiber Network Design – Version 1.4

**Author:** Ruben Brionez Jr
**Contributors:** Tiffany Rufo, Omni Fiber GIS Department
**Version:** 1.4 (July 2025)

---

## 🔍 Overview

This repository contains a modular, class-based Python script for processing and exporting Bill of Materials (BOM) data from fiber network designs within FDH (Fiber Distribution Hub) boundaries. Built for ArcGIS Pro environments, the tool integrates with ArcGIS Online (Portal) feature layers to extract, process, and summarize network components and export the results to a pre-formatted Excel template.

## 🚀 Major Updates in v1.4

* Refactored to object-oriented architecture using the `BOMProcessor` class
* Added robust error handling and ArcPy messaging
* Enhanced field and geometry validation
* Fully decoupled FDH selection logic for interactive or parameter-based use
* Modular export integration using `openpyxl`
* Extended logging and debugging output

## 🧰 Key Features

* Supports spatial and attribute selection from ArcGIS Portal-hosted layers
* Processes the following network components:

  * **Conduit**: UG1/UG2 footage, coupler logic
  * **Cable**: SP1/SP2/SP3, fiber counts, splice slack, maintenance slack
  * **Strand**: with sag and anchor estimations
  * **Structures**: flowerpots, vaults, risers, cabinets (passive/active)
  * **Slackloops**: detailed footage and splice buffer analysis
  * **Address Points**: filtered by FDH, MDU, and Do Not Build overlays
  * **Poles and Guys**: count and location analysis
* Calculates:

  * Fiber and strand footage totals
  * Fiber miles (F1/F2)
  * Slack loop and coupler material requirements
  * Drops, slack, lashing wire, snowshoes
* Exports to a structured BOM Excel template with precise cell mappings

## 📂 File Structure

```
project_root/
├── BOM_Processing_v1.4.py       # Main script with BOMProcessor class
├── TEST - BOM Template_03052025.xlsx
└── README.md
```

## 📈 Output

* Exports Excel BOM to:

  * `RateCard` and `RateCard_E`: Vendor-specific rate info
  * `Summary`: Overall design totals (addressable, fiber miles, slack)
  * `Engineering`: Engineering-specific breakdowns of fiber/cable/conduit

## ⚙️ Requirements

* ArcGIS Pro (Python 3.x environment)
* ArcPy (included with ArcGIS Pro)
* ArcGIS API for Python (`arcgis`)
* `openpyxl`

Install missing packages with:

```bash
conda install -c esri arcgis
pip install openpyxl
```

## 🗂 Required Portal Feature Layers

Ensure the following layers are accessible and configured in the script:

* FDH\_Boundary
* Conduit, Cable, Strand
* Structures, Poles, Risers
* Cabinets (Passive/Active)
* Slackloops, Guys, Drops
* MDU and DNB Boundaries
* Address Points

## ▶️ How to Use

1. Open ArcGIS Pro and load your project.
2. Add the relevant Portal layers to the map or configure them in script parameters.
3. Run `BOM_Processing_v1.4.py` from the ArcGIS Python window or as a script tool.
4. Provide parameters:

   * `cab_id`: FDH ID to process
   * `Run Export`: `True` or `False`
   * Vendor rates for design and construction
   * Output Excel path (optional)
5. Review ArcGIS messages and resulting Excel file.

## 🧪 Example Output Variables

Key calculated outputs include:

* `total_ug1ft`, `total_strand_material_ftg`, `fiber_144`, `sp1_total`, `slack_total_ft`, `fp_count`, `ae_vs_ug_ratio`, `total_f2_miles`, `vault_total`, `drop_total_footage`

## 📊 Sample Excel Output

| Metric           | Value      |
| ---------------- | ---------- |
| UG1 FT           | 1,258.75   |
| F1 Miles         | 4.92       |
| Slack (Total)    | 820.55     |
| SP1/2/3 Combined | 422.00     |
| Drops (Footage)  | 2,152.33   |
| Anchors & Guys   | 18 anchors |

## 🚧 Future Improvements

* Support for multi-FDH batch processing
* Integrate Slackloop photos into output
* Add config file for layer mappings and Excel cell positions
* Automate vendor selection from Portal metadata
* Export as CSV or GeoJSON for downstream ETL workflows

## 📜 License

Proprietary – for internal use by Omni Fiber only.

## 📞 Contact

For support or feedback, reach out via Omni Fiber GIS or Engineering Teams.
