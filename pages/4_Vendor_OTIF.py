


# Colab OTIF report script (auto-detect .xls/.xlsx, normalize columns, same logic as before)
# Run this cell in Google Colab. When prompted, upload your Excel file.

# Install helpful packages (quiet)
!pip install --quiet openpyxl xlrd reportlab

import io
import os
import re
import sys
from datetime import datetime
import pandas as pd
import numpy as np

# Reportlab imports for PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# ---------- CONFIG / MAPPINGS ----------
# Placeholder item-category mapping (extend with your real map)
# --------- item map (placeholder) ----------

# ---------- CONFIG / MAPPINGS (same as your Streamlit app) ----------
df1 = pd.DataFrame({'Material Code': ['4AO005', '1DAT04S', '1DCT01', '2AE06', '2CC02', '4BT021G', '2AB01-C', '4BT008G', '4BT008G', '4BT008G', '4BT011G', '4BT015G', '4BT015G', '4BT019G', '4BT031G', '4BT035G', '4BT036G', '4BT036G', '4BT038G', '4BT050G', '4BT050G', '1FCM01', '2CD09-C', '5DAGB01', '1DCM03', '1FCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '2CD08', '3BG02R', '2AB02', '1ICM01Y', '2AB01-C', '2AB01-C', '2CA07-C', '3CG02', '3CG02', '4CT045G', '2AD01-C', '12OT07', '2CA07-P', '4BO040G', '1FCM03', '4BT017G', '4BT01G', '4BT043G', '1BAM01', '1BAM01', '1BAM01', '1BAM01', '2CC09-C', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '4BT036G', '4BT036G', '1DCT01R', '4BO015G', '4BO015G', '1DCT01', '4BO018G', '2CB01-C', '1BAM01', '1DCT01R', '3BG02R', '4BO015G', '4BO015G', '1BCM03', '1DAT04S', '1FAM01', '3BG02', '4BO040G', '4BO040G', '4BT013G', '4BT051G', '1DAT01L', '2AB02', '2AB02', '2AB02', '2CC09-C', '4BO049G', '4BO049G', '4BT003G', '4BT010G', '4BT050G', '2AA01-C', '1DAM01', '1DCT01R', '1FCT01R', '2CA23-C', '2CA23-C', '4BT025G', '1FAM01', '3BS02', '3BS02', '4BO040G', '4BO040G', '7PBA01-A', '7PBA01-A', '7PBB01-B', '7PBB01-B', '7PBB01-C', '4BT011G', '1DAT01L', '1DCT01', '2AA01', '3BG02-U', '4BT025G', '4BT025G', '4BT049G', '4BT049G', '4BT015G', '1FCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT011G', '4BT035G', '4BT036G', '4BT036G', '1HCM01', '4BT019G', '1CCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '4BT006G', '5DAGB01', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '1DCT01', '2CE02', '2CE02', '2CE02', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '4BT022G', '4BT031G', '4BT035G', '4BT036G', '4BT036G', '4BT036G', '4BT036G', '1HCM01', '1BCM03', '1ICM01', '4BO043G', '4BT034G', '4BT034G', '4BT034G', '2AB07', '4BT008G', '1DAM01', '4AO002', '4BT025G', '1ACT01', '1DCT01R', '1FCT01R', '3AG02', '3BG02R', '4BO015G', '1BCM03', '1DAT01L', '1HCM01', '2AB02', '2AB02', '2AB02', '4BO018G', '4BO01G', '4BO01G', '4BO049', '4BT025G', '2AD02', '4BO049G', '4BO049G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT019G', '1FCM01', '1FCM01', '1FCM01', '1FCM01', '2AD02', '2AD02', '2AA02', '3BG02', '1DAT01L', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1BCT01S', '1HCM01', '3BG02', '3BG02', '3BG02', '3BG02', '4BO01G', '4BO049G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT050G', '1FCM01', '4BO044G', '5DAGB01', '4BT063G', '4BT008G', '1BAM01', '1HAM01', '2AE06', '2CC09-C', '4BO015G', '4BT015G', '4BT050G', '2CA07', '4BT013G', '4BT013G', '4BT038G', '1BCT01S', '1BCT01S', '2AA02', '4BO01', '1ECM03', '4BT006G', '1BCM03', '2CC09-C', '2CE04', '4BT034G', '4BT034G', '1DAM01', '1FCM01', '2CC09', '2CC09', '4BO031G', '4BO044G', '1CCM03', '1CCM03', '2AE02', '1DCT01R', '1FCT01R', '2AC07', '2CB01-C', '4BO015G', '1DCM03', '1DCM03', '1DCM03', '1DAM03', '1ICM01', '2CD08', '2CD08', '4BO01G', '4BO044G', '4BT025G', '1ECM03', '4BT008G', '4BT011G', '4BT045G', '1BCT01S', '1BCT01S', '1BCT01S', '1DAT01L', '2AB07', '4BT021G', '4BT02G', '7PBB04-A', '7PBB04-B', '7PBB04-C', '2CD09-C', '1HCM03', '4BT036G', '1DAM01', '1DAT04S', '1HCM01', '2AC07', '2CE02', '4BO015G', '4BO031G', '4BO044G', '4BT021G', '4BT021G', '4BT021G', '1DAM03', '4BO049G', '4BT025G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1ACT01', '3AG04', '3AG04', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4AT003G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT021G', '4BT008G', '4BT011G', '1FCM01', '2AD02', '4BT015G', '5DAGB01', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '1DCM01', '2CD08', '4BT024G', '1DCM03', '4BT036G', '4BT063G', '2CB01-C', '2CD09-C', '2CE04', '4BT008G', '4BT011G', '4BT036G', '4BT038G', '4BT050G', '1BCM03', '4BT011G', '1DCM03', '5CAGB01', '6C02', '2AB01-C', '2AD02', '2AA01-C', '4BO015G', '1CCM03', '4BT008G', '4BT008G', '4BT022G', '4BT015G', '4AO005G', '4BT034G', '4BT016G', '1BAM01', '1ECT01R', '1HCM01', '2AA01-C', '2CC09', '4BO044G', '4BT015G', '4BT021G', '7PBC01-A', '4BT006G', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1DAT04S', '1DAT04S', '1FCM01', '2CE10-C', '4BO044G', '4BT011G', '4BT033G', '1BAM01', '1DCT01', '4BT02G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1CCM03', '1DCM03', '1FCM01', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1HCM01', '1HCM01', '1HCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1ACT01', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1DAM01', '1DCM03', '1FCM01', '1HCM01', '1HCM01', '1ICM01', '2AA01-C', '2AB01-C', '2AB02', '2AB02', '2AD02', '2AD02', '2AD02', '2AD02', '2AE02', '2CA07-C', '2CB01-C', '2CC09-C', '2CD08', '2CD08', '2CD09-C', '2CD09-C', '2CD09-C', '2CD09-C', '2CE02', '2CE02', '2CE10-C', '3AG02', '3AG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3CG02', '3CG02', '3CG02', '3CG02', '4BT02G', '4BT02G', '1FCT02', '1HCM03', '3BG06', '3BG06', '3BG06', '3BG06', '3BG06', '3BG06', '4BT015G', '4BT065', '4BT065', '4BT066', '4BT066', '4BT021G', '4BT011G', '4BT013G', '4BT033G', '4BT033G', '4BT033G', '4BT008G', '4BT008G', '2AB07', '2AE06', '2CD09-C', '2CD09-C', '2CD09-C', '2CD09-C', '4BT033G', '4BT033G', '4BT033G', '1FCM01', '4BO044G', '4BT006G', '4BT008G', '4BT008G', '1DAM03', '1DCT01R', '1FCT01R', '4BO015G', '4BO015G', '4BT034G', '4BT034G', '4BT034G', '7PBB04-A', '7PBB04-B', '7PBB04-C', '4BT019G', '4BT02G', '4BT02G', '7PBB04-A', '7PBB04-B', '7PBB04-C', '4BT003G', '4BT011G', '4BT013G', '4BT019G', '4BT032G', '4BT051G', '1DCM01', '4BO015G', '4BO015G', '4BO01G', '4BO01G', '4BO01G', '4BO044G', '4BT011G', '4BT011G', '4BT015G', '4BT015G', '4BT021G', '4BT021G', '4BT021G', '4BT025G', '4BT025G', '4BT025G', '4BT02G', '4BT02G', '4BT033G', '4BT033G', '4BT034G', '4BT034G', '4BT034G', '1BAM01', '4BT034G', '4BT034G', '4BT034G', '2CA07-C', '1FAM01', '4BO040G', '4BO040G', '4BT008G', '4BT008G', '4BT011G', '2AB02', '2AB02', '2AB02', '2AB02', '5DAGB01', '7PBB01-A', '4BT009G', '2CC09', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '1DCM03', '2AD02', '2AD02', '2CE10-C', '3BG05', '4BT019G', '1BCM03', '1HAM01', '4BO018G', '4BO055G', '7PBB01-A', '7PBB01-B', '7PBB01-C', '4BT015G', '3BG04', '4BT008G', '4BT008G', '4BT008G', '9PFSA06-A', '9PFSA06-B', '9PFSA06-C', '9PFSA06-D', '1HCM03', '1HCM03', '4BT011G', '4BT011G', '4BT013G', '1DAM01', '1IAM01', '1ICM01Y', '2AB07', '2CD07', '3CG02', '3CG02', '3CG02', '3CG02', '4BO058G', '4CT002G', '4CT002G', '5CAGB01', '5DAGB01', '6C02', '7PBB04-A', '7PBB04-B', '7PBB04-C', '8CP06', '9PL01', '2AD02', '2CE02', '4BT021G', '4BT021G', '4BT021G', '4BT022G', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1DCT01', '1HCM01', '2AA01-C', '2AB02', '2AB02', '2AD01-C', '4AO043G', '1DCM03', '4BT015G', '4BT015G', '1DAT01L', '1DCM03', '2AB02', '2AB02', '4AO001G', '1FCM03', '4BT008G', '4BT008G', '4BT008G', '4BT036G', '1FCM01', '2AE06', '2AE06', '1CCM03', '4BT063G', '1FAM01', '1BAM01', '2AA01-C', '1DCM03', '1BCM03', '1HCM01', '2AA01-C', '2AB01-C', '2AD02', '2AD02', '2CA02', '2CB01-C', '2CD09-C', '4BT034G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '1ACT01', '1ACT01', '3AG02', '3AG02', '3AG02', '3AG02', '4AO005G', '4AO005G', '1FAT01', '2CE08-C', '4BT008G', '7PBB04-A', '7PBB04-B', '2AA01-C', '2AB01-C', '2CB01-C', '2CD09-C', '1BAT01', '2CB01-C', '4BT063G', '1DCT01R', '1FCT01R', '2AD01-C', '3BG02', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1FCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '1FCM01', '1FCM01', '1FCM01', '1FCM03', '2AD02', '2AD02', '2CA07-C', '2CA27', '2CC02', '4AO001G', '4BO01G', '4BO01G', '4BT011G', '4BT025G', '1BAM01', '4BO034G', '4BO034G', '4BT068G', '2CC09-C', '2CD02', '2CD02', '2CD02', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1DAM03', '1DCM01', '4BT008G', '4BT008G', '4BT008G', '4BT011G', '4BT013G', '4BT015G', '4BT022G', '4BT022G', '4BT022G', '4BT036G', '4BT036G', '4BT036G', '4BT050G', '4BT063G', '2CA07-C', '2CB01-C', '1DCM03', '2AE02', '1FCM03', '4AO005G', '4AO005G', '7PBB01-B', '7PBB01-C', '1FCM03', '1HCM03', '4BT008G', '4BT011G', '4BT011G', '4BT011G', '4BT013G', '4BT015G', '4BT019G', '4BT043G', '1BAM01', '4BT011G', '4BT036G', '4BT050G', '1DCM03', '1DCM03', '1DCM03', '1FCM03', '1FCM03', '1BCM03', '1DAM03', '2CE02', '2CE02', '2CE02', '2CE02', '4BT005G', '4BT034G', '4BT034G', '4BT034G', '1FCM01', '4BO044G', '4BT011G', '4BT050G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1FAM01', '1FAM01', '1FCM01', '2AD02', '2AD02', '2CA07', '2CE10-C', '3BS02', '2CE02', '2CE02', '2CE02', '2CE02', '4BT006G', '1DCT01', '2AB02', '2AD02', '2AD02', '2AD02', '3AC02', '3BG04', '4BT034G', '4BT034G', '4BT034G', '2CB01', '4BT008G', '4BT008G', '4BT063G', '2AB01-C', '5DAGB01', '5DAGB01', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT008G', '4BT016G', '4BT036G', '1FCM03', '1FCM03', '1HCM03', '3BG07', '3BG07', '1BAM01', '2CC09-C', '2CE02', '2CE02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1DAM03', '1DCT01', '2AE06', '4AO001G', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1HAM01', '2AA01-C', '2CA07-C', '4BO055G', '1DAT01L', '4BO049G', '1ACT01', '1BCM03', '4AO002G', '4BT034G', '4BT034G', '4BT034G', '4BT034G', '1BCT01', '1DCM03', '2CC09', '2CE08-C', '4BT011G', '4BT011G', '4BT049G', '4BT003G', '4BT032G', '4BT050G', '1BAM01', '1DAM01', '2CD02', '2CD02-LE', '4BO031G', '7PBB04-A', '7PBB04-B', '4BT048G', '1DAM03', '2CC09-C', '2CD02-LE', '2CD02-LE', '2CE08-C', '4AO043G', '4AO043G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '4BT017G', '4BT017G', '4BT01G', '4BT022G', '4BT022G', '4BT022G', '1DCM01', '2AB07', '2CE10-C', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '7PBB04-A', '7PBB04-B', '1BAM01', '2CC09', '2CD04', '4BT024G', '4BT02G', '7PBA01-A', '7PBB01-B', '7PBB01-B', '7PBB01-C', '4BT008G', '4BT008G', '4BT008G', '4BT021G', '2AB01-C', '2CD02', '2CD02', '1DCM03', '12OT11', '12OT11', '12OT11', '12OT11', '12OT11', '12OT11', '2CA07-C', '2CB01-C', '2CD02', '2CD02', '2CD02', '2CD02', '4BO036', '1HAM01', '2AB01-C', '2CE10', '4BO055G', '4BO01G', '4BO01G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT015G', '4BT015G', '4BT015G', '4BT022G', '4BT022G', '2CC09-C', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '2CB01-C', '1DCT01-PB', '1JCT01-PB', '2CB34-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '3BG02-PB', '4BO036-PB', '4BO043-PB', '4BO061-PB', '4BO063-PB', '4BT011G', '4BT019G', '4BT019G', '2AB07', '2AE02', '4BT006G', '4BT009G', '4BT009G', '4BT012G', '1DCT01R', '1BAM01', '4BO034G', '4BT025G', '2CC09-C', '1DCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT015G', '4BT050G', '4BT036G', '4BT036G', '4BT036G', '4BT025G', '4BT025G', '4BT025G', '4BT025G', '2CD02', '4BT022G', '2AB02', '2AB02', '2AD02', '2AD02', '4BT021G', '3AG02', '1FAM01', '2CA07-C', '4BO015G', '4BO015G', '2AA01-C', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1FCM03', '4BO013G', '4BT022G', '4BT022G', '2CD02', '2CD02', '2CD02', '4BT036G', '4BT036G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '2AD08', '3BS02', '1FCM01', '4BO044G', '3BG08', '3BG08', '5DAGB01', '6C01', '6C02', '1DCM03', '1DCM03', '1HCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT015G', '1DAM03', '1DAM03', '2CA23-C', '7PBB04-A', '7PBB04-B', '1BCT01-G', '1DCT01-G', '3BG04-G', '4BT019-G', '5CAGB01', '4BT036G', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1FCM01', '2AA01-C', '4BO044G', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1DCM03', '1FCM03', '1DCM01', '1DCM03', '4BT008G', '4BT008G', '4BT008G', '4BT011G', '4BT015G', '4BT015G', '2CA07', '2CC02', '2CC09', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '9PFSA03-A', '2AA02', '4BT011G', '4BT013G', '4BT063G', '2CD08-EP', '4BO044G', '1ICM01-G', '2CC09-C', '1BCT01', '4BO018G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '7PBB01-C', '1BCM03', '2AA01-C', '4BT034G', '4BT034G', '4BT034G', '4BT021G', '1CCM03', '2AC07', '1BCM03', '1BCM03', '1DAM03', '4BT034G', '1BCM03', '2CC09', '2CC09', '1HCM03', '4BT008G', '4BT008G', '2CE02', '2CE02', '2CE02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-A', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '7PBB04-B', '1DCM03', '1DCM03', '4BT006G', '4BT008G', '4BT008G', '4BT02G', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '2CE02', '4BT02G', '1DCM03', '1DCM03', '4BT011G', '4BT011G', '4BT011G', '4BT050G', '2CA07-C', '4AO004G', '4BT006G', '4BT006G', '1DAM01', '1DCM03', '1FCM03', '1HAM01', '1HCM01', '1ICM01', '1ICM01', '2CB01', '2CC09-C', '4BO015G', '4BO055G', '4BT011G', '4BT025G', '4BT025G', '1DCM04', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '1FAM01', '1FAM01', '2AD01-C', '2AD01-C', '4BO040G', '4BO036-PB', '4BO043-PB', '4BO061-PB', '4BO063-PB', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '2CE10', '2CE10', '4BT015G', '2AA02', '4BO031G', '4BT034G', '4BT034G', '4BT034G', '4BT034G', '1BCM03', '1DCT01R', '1ECT01R', '2AD01-C', '2AD01-C', '7PBB01-A', '7PBB01-B', '7PBB01-C', '4AO004', '1DCM01', '1FAM01', '4BT021G', '4BT021G', '4BT021G', '4BT021G', '4BT024G', '1FCM03', '4BT013G', '4BT013G', '4BT013G', '4BT015G', '1BAM01', '4BT02G', '4BT02G', '1BCM03', '1DCM03', '1DCM03', '1ACT01', '3AG02', '3AG02', '3AG02', '3AG02', '4AO001G', '4AO002G', '4BT006G', '6C02', '2CA07', '1BCT01X', '2CE10', '2CE10', '1IAM01', '2AE02', '4BO058G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '4BT008G', '1BCM03', '1BCM03', '1DAT01S', '3BG02', '3BG02', '3BG02', '3BG02', '4BT025G', '4BT025G', '4BT034G', '1HCM03', '4BT015G', '1BAM01', '2AA01-C', '2AD01-C', '2AD01-C', '2CA07-C', '1BCM03', '1BCM03', '1BCM03', '1BCM03', '2CB01-C', '2CB01-C', '2CD08', '4BT011G', '2AA02', '2AD02', '2AD02', '2AD02', '8CP06', '9PFSA01-B', '9PFSA03-A', '9PL01', '1DCT01', '2CC09', '1FCM01', '1FCM01', '1BAM01', '2CC09-C', '2CC09-C', '2CC09-C', '4BT021G', '4BT021G', '1DAT01L', '1DCM03', '4BT021G', '4BT021G', '4BT021G', '5DAGB01', '4BT017G', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1BCT01S', '1FCM03', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '4BT02G', '7PBB02-A', '7PBB02-B', '7PBB02-C', '2AB02', '2AB02', '1DAM01', '1DAM01', '4BT034G', '4BT034G', '2AD01-C', '1DCM03', '1DCM03', '1DCT01R', '1FCT01R', '2CE10', '2CE10', '3BG04', '4BO015G', '1BCM03', '1BCM03', '1DCM03', '1DCM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT011G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '4BT033G', '2AB05', '2AB05', '1FCM01', '1FCM03', '1FCM03', '1FCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '1HCM03', '2AA01-C', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT033G', '4BT02G', '4BT048G', '2AD01-C', '2CC02', '7PBB02-A', '7PBB02-A', '7PBB02-A', '7PBB02-B', '7PBB02-B', '7PBB02-B', '7PBB02-C', '7PBB02-C', '4BT011G', '2AB35', '2AB35', '2CA07', '2CA07', '2CB01-C', '2CD04', '1DAM03', '4BO040G', '2CA07', '3AC02', '4AO044G', '4BT036G', '1FAM01', '2AB07', '5CAGB01', '2AA01-C', '3BG10', '4BT065G', '4BT066G', '1BAM01', '1DAT01L', '1DCT01R', '2AB01-C', '2AB01-C', '2AD02', '1DCM03', '1ECM03', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '4BT017G', '2CD02', '2CD02', '2CD02', '1CCM03', '10CTC01', '1FCT01R', '1FCT01R', '2AA02', '4BT008G', '2AD02', '4BT050G', '4BT034G', '1BAM01', '2AB01-C', '4BT015G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1HCM03', '1HCM03', '4BT008G', '4BT008G', '4BT008G', '4BT008G', '4BT036G', '1CCM03', '2CD02', '4BT011G', '4BT013G', '4BT050G', '2CD08', '1FCM01', '4AO001G', '4BO015G', '4BO044G', '5CAGB01', '6C01', '7PBA01-A', '7PBB01-B', '7PBB01-C', '8CP06', '9PL01', '6C02', '3BG02', '3BG02', '3BG02', '3BG02', '3BG02', '1BCM03', '1DAT01L', '2AA01-C', '4BT033G', '2CC09', '1DCM03', '2AD04-K', '4BT015G', '4BT035G', '1FCM01', '1FCM01', '1FCM01', '4BO044G', '1DCT01', '1DCT01R', '1BCT01X', '4AT002G', '4BT02G', '1BCT01S', '1FCM01', '2CD02', '9PFSA01-B', '9PFSA01-B', '9PFSA03-A', '9PFSA03-A', '3BG02', '3BG02', '3BG02', '9PCCC01-NA', '9PCCC02-NA', '1FCM03', '4BT063G', '7PBA01-A', '7PBB01-B', '7PBB01-C', '1DCT01R', '1FCT01R', '2CB01-C', '4BO015G', '1DAM03', '4BO031G', '1HCM03', '4BT015G', '4BT036G', '4BT034G', '1FAM01', '9PCCC01-NA', '9PCCC02-NA'], 'Item Category': ['Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Cap', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Ampoule', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Ampoule', 'Al Tube', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Plastic Bottles', 'Plastic Bottles', 'Cap', 'Cap', 'Plastic Nozzle', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Cap', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Vial', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Vial', 'Seal', 'Cap', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Outer cap', 'Ampoule', 'Vial', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Cap', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Vial', 'Cap', 'Collar', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Outer cap', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Outer cap', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Cap', 'Plastic Bottles', 'Seal', 'Ampoule', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Plunger Rod', 'Plunger Rod', 'Plunger Rod', 'Plunger Rod', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Cap', 'Cap', 'Collar', 'Plastic Bottles', 'Cap', 'Plunger Rod', 'Cap', 'U plug', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Vial', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Vial', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Cap', 'Cap', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Plastic Bottles', 'Cap', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Cap', 'Plastic Nozzle', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Al Tube', 'Al Tube', 'Al Tube', 'Al Tube', 'Al Tube', 'Al Tube', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Ampoule', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Rubber Stopper', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Ampoule', 'Rubber Stopper', 'Vial', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Cap', 'Collar', 'Collar', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Plastic Bottles', 'Cap', 'Vial', 'Vial', 'Rubber Stopper', 'Seal', 'Cap', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'PFS Syringe', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Vial', 'Ampoule', 'Vial', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Plastic Nozzle', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Cap', 'Cap', 'Cap', 'Cap', 'Cap', 'Cap', 'Cap', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Collar', 'Ampoule', 'Vial', 'Ampoule', 'Ampoule', 'Vial', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Seal', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Vial', 'Seal', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Cap', 'Plunger Stopper', 'PFS Syringe', 'U plug', 'Vial', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Cap', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Rubber Stopper', 'Seal', 'Vial', 'Vial', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Ampoule', 'Plastic Bottles', 'Plastic Bottles', 'Plastic Bottles', 'Cap', 'Cap', 'Cap', 'Plastic Nozzle', 'Plastic Nozzle', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Seal', 'Ampoule', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Cap', 'Ampoule', 'Rubber Stopper', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Vial', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Seal', 'Ampoule', 'Ampoule', 'Ampoule', 'Vial', 'Plunger Rod', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Seal', 'Seal', 'Seal', 'Seal', 'Seal', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Seal', 'Ampoule', 'Vial', 'Seal', 'Seal', 'Seal', 'Cap', 'Collar', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Cap', 'U plug', 'Collar', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Ampoule', 'Vial', 'Ampoule', 'Seal', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Vial', 'Vial', 'Vial', 'Seal', 'Seal', 'Vial', 'Vial', 'Ampoule', 'Plunger Stopper', 'Plunger Stopper', 'PFS Syringe', 'PFS Syringe', 'Rubber Stopper', 'Rubber Stopper', 'Rubber Stopper', 'Plunger Rod', 'Rubber Stopper', 'Vial', 'Seal', 'Plastic Bottles', 'Cap', 'Plastic Nozzle', 'Vial', 'Vial', 'Ampoule', 'Seal', 'Vial', 'Seal', 'Vial', 'Seal', 'Seal', 'Seal', 'Vial', 'Plunger Rod', 'Rubber Stopper']})


# Canonical column names used by the processing logic
COL_MAT_TYPE = 'Mat Type'
COL_MATERIAL_CODE = 'Material Code'
COL_MATERIAL_NAME = 'Material Name'
COL_UOM = 'UOM'
COL_PO_DT = 'P.O. Dt.'        # canonical P.O. date (with dot)
COL_PO_NO = 'P. O. No.'      # canonical PO number
COL_SUPPLIER = 'Supplier'
COL_PO_QTY = 'PO Qty.'       # canonical PO qty
COL_GNR_DT = 'GNR Dt.'       # canonical GRN date
COL_INWARD_QTY = 'Inward Qty.'
COL_ITEM_CAT = 'Item Category'  # optional

REQUIRED_COLS = [
    COL_MAT_TYPE, COL_MATERIAL_CODE, COL_MATERIAL_NAME, COL_UOM,
    COL_PO_DT, COL_PO_NO, COL_SUPPLIER, COL_PO_QTY, COL_GNR_DT, COL_INWARD_QTY
]

DEFAULT_RULES = {"RM": 30, "SPM": 15, "TPM": 15}
DEFAULT_PPM_LT = 30

PPM_CATEGORY_MAP = {
    7:  ['Vial', 'Rubber Stopper', 'Rubber', 'Stopper', 'Seal', 'Cap', 'Collar', 'Inner Cap', 'Outer Cap'],
    12: ['Ampoule', 'Amp'],
    90: ['Pfs Syringe', 'Plunger Stopper', 'Plunger', 'U plug', 'U-plug'],
    15: ['Al Tube', 'Plastic Bottle', 'Plastic Nozzle', 'Nozzle'],
}

# Colab-specific behavior: default lead time assigned to unknown Mat Types (change if needed)
DEFAULT_UNKNOWN_LEAD_TIME = 30   # days assigned if Mat Type is unknown
# Optional override dict (set before uploading if you wish)
# Example: custom_lead_times = {"CUSTOMTYPE": 20}
custom_lead_times = {}

# ---------- UTILITIES: robust Excel reading & column normalization ----------
def read_excel_auto(path_or_buffer):
    """
    Read an excel file (path or file-like) while auto-selecting the engine based on file extension.
    Works with .xls (xlrd) and .xlsx/.xlsm (openpyxl). Falls back if necessary.
    """
    # If a path-like string was passed, we can inspect the extension
    engine = None
    fname = None

    # If the argument is a path string
    if isinstance(path_or_buffer, str):
        fname = path_or_buffer
        ext = os.path.splitext(fname)[1].lower()
        if ext == '.xls':
            engine = 'xlrd'
        elif ext in ('.xlsx', '.xlsm', '.xltx', '.xltm'):
            engine = 'openpyxl'
        else:
            # unknown extension: try openpyxl then xlrd
            engine = None

        if engine:
            try:
                df = pd.read_excel(fname, engine=engine)
                print(f"Read '{fname}' using engine='{engine}'.")
                return df
            except Exception as e:
                print(f"Failed reading with engine='{engine}': {e}. Will attempt fallbacks.")
                # fallthrough to generic attempts

    # If argument is file-like (e.g., BytesIO), try openpyxl first, then xlrd
    attempts = []
    try:
        # prefer openpyxl for modern xlsx files
        df = pd.read_excel(path_or_buffer, engine='openpyxl')
        print("Read Excel using engine='openpyxl'.")
        return df
    except Exception as e_open:
        attempts.append(("openpyxl", str(e_open)))
        try:
            df = pd.read_excel(path_or_buffer, engine='xlrd')
            print("Read Excel using engine='xlrd'.")
            return df
        except Exception as e_xl:
            attempts.append(("xlrd", str(e_xl)))
            # If both fail, surface a helpful message
            msg = "Failed to read Excel file. Attempts:\n"
            for eng, err in attempts:
                msg += f" - engine={eng}: {err}\n"
            raise ValueError(msg)

def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Robustly map many variants of column names to canonical names used by the script.
    Preserves any columns not recognized (they remain with near-original name, cleaned).
    """
    def map_one(col):
        orig = str(col).strip()
        # clean spaces and NBSP
        s = re.sub(r'\u00A0', ' ', orig)
        s = re.sub(r'\s+', ' ', s).strip()
        key = re.sub(r'[\s\.\-_/()]', '', s).lower()

        # mapping dictionary based on common variants
        if key in ('mattype','materialtype'):
            return COL_MAT_TYPE
        if key in ('materialcode','matcode','material_code'):
            return COL_MATERIAL_CODE
        if key in ('materialname','itemname','material'):
            return COL_MATERIAL_NAME
        if key == 'uom':
            return COL_UOM
        if key in ('podt','podat','podate','podatet','podt'):
            return COL_PO_DT
        if key in ('pono','ponumber','po'):
            return COL_PO_NO
        if key in ('supplier','suppliername'):
            return COL_SUPPLIER
        if key in ('poqty','poqty','purchaseorderqty','quantityordered'):
            return COL_PO_QTY
        if key in ('gnrdt','grndt','grndate','grn','grndate'):
            return COL_GNR_DT
        if key in ('inwardqty','inwardquantity','receivedqty','receivedquantity'):
            return COL_INWARD_QTY
        if key in ('itemcategory','itemcat','category'):
            return COL_ITEM_CAT
        # else return the cleaned original (preserve punctuation if user wants)
        return orig

    new_cols = [map_one(c) for c in df.columns]
    df = df.copy()
    df.columns = new_cols
    return df

# ---------- HELPERS: original logic, but using canonical column variables ----------
def compute_lead_time_for_row(row: pd.Series, rules: dict):
    mat_type = str(row.get(COL_MAT_TYPE, "")).strip().upper()
    if mat_type in rules:
        return rules[mat_type]
    if mat_type == "PPM":
        item_cat = str(row.get(COL_ITEM_CAT, "") or "").strip()
        if item_cat:
            low = item_cat.lower()
            for lt, cats in PPM_CATEGORY_MAP.items():
                if low in [c.lower() for c in cats]:
                    return lt
        return DEFAULT_PPM_LT
    return np.nan

def ensure_types_and_drop_nulls(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df[COL_PO_DT] = pd.to_datetime(df[COL_PO_DT], errors="coerce", dayfirst=True)
    df[COL_GNR_DT] = pd.to_datetime(df[COL_GNR_DT], errors="coerce", dayfirst=True)
    df[COL_PO_QTY] = pd.to_numeric(df[COL_PO_QTY], errors="coerce")
    df[COL_INWARD_QTY] = pd.to_numeric(df[COL_INWARD_QTY], errors="coerce")
    df = df.dropna(subset=[COL_PO_DT, COL_GNR_DT, COL_PO_QTY, COL_INWARD_QTY]).copy()
    return df

def merge_item_category(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    if COL_ITEM_CAT not in df.columns:
        df = df.merge(df1, on=COL_MATERIAL_CODE, how="left")
    else:
        df = df.merge(df1, on=COL_MATERIAL_CODE, how="left", suffixes=("", "_map"))
        if "Item Category_map" in df.columns:
            df[COL_ITEM_CAT] = df[COL_ITEM_CAT].fillna(df["Item Category_map"])
            df.drop(columns=["Item Category_map"], inplace=True, errors="ignore")
    df[COL_ITEM_CAT] = df.get(COL_ITEM_CAT, "").fillna("")
    return df

def compute_po_level_metrics(df: pd.DataFrame):
    # Product-level fulfillment (sum duplicates)
    df_po_item = (
        df.groupby([COL_PO_NO, COL_MATERIAL_CODE], as_index=False)
          .agg({COL_PO_QTY: "sum", COL_INWARD_QTY: "sum"})
    )
    df_po_item["Fulfilled"] = (df_po_item[COL_INWARD_QTY] >= 0.95 * df_po_item[COL_PO_QTY]).astype(int)

    # PO-level In-Full
    df_po_status = (
        df_po_item.groupby(COL_PO_NO)["Fulfilled"]
                  .min()
                  .reset_index()
                  .rename(columns={"Fulfilled": "PO_Fulfilled"})
    )
    df_line = df.merge(df_po_status, on=COL_PO_NO, how="left")

    # PO-level On-Time
    def po_ontime(group: pd.DataFrame) -> int:
        due_dates = group[COL_PO_DT] + pd.to_timedelta(group["Lead Time"], unit="D")
        return int((group[COL_GNR_DT] <= due_dates).all())

    po_ontime_df = df_line.groupby(COL_PO_NO).apply(po_ontime).reset_index(name="OnTime")
    df_line = df_line.merge(po_ontime_df, on=COL_PO_NO, how="left")

    # Collapse to one-row-per-PO (keep other PO-level columns)
    df_po = df_line.drop_duplicates(subset=[COL_PO_NO]).copy()
    df_po["OTIF"] = (df_po["PO_Fulfilled"].astype(int) * df_po["OnTime"].astype(int)).astype(int)

    # Use LAST GNR date per PO for bucketing
    po_last_grn = (
        df.groupby(COL_PO_NO, as_index=False)[COL_GNR_DT]
          .max()
          .rename(columns={COL_GNR_DT: "PO_GNR_Dt"})
    )
    df_po = df_po.merge(po_last_grn, on=COL_PO_NO, how="left")
    df_po[COL_GNR_DT] = pd.to_datetime(df_po["PO_GNR_Dt"])
    df_po.drop(columns=["PO_GNR_Dt"], inplace=True)

    # Add Year/Month parts
    df_po["Year"] = df_po[COL_GNR_DT].dt.year
    df_po["MonthNum"] = df_po[COL_GNR_DT].dt.month
    df_po["Month"] = df_po[COL_GNR_DT].dt.strftime("%b")

    return df_line, df_po

def generate_failed_orders_pdf_colab(breaches_df: pd.DataFrame, vendor_stats: pd.DataFrame, year: int, output_path: str):
    buf = io.BytesIO()
    width, height = A4
    c = canvas.Canvas(buf, pagesize=A4)
    margin_x = 20 * mm
    y = height - 20 * mm
    line_height = 8 * mm

    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin_x, y, f"ALL OTIF FAILED ORDERS — {year}")
    y -= 12 * mm

    vendor_order = vendor_stats[vendor_stats["OTIF_Failures"]>0].sort_values("OTIF_Failures", ascending=False)

    vendor_idx = 1
    for _, vr in vendor_order.iterrows():
        vendor = vr["Supplier"]
        failures = int(vr["OTIF_Failures"])
        total_orders = int(vr["Total_Orders"])
        otif_pct = float(vr["Vendor_OTIF_pct"])
        contrib_pct = float(vr["Total_Contribution_pct"])

        vendor_group = breaches_df[breaches_df["Supplier"] == vendor].sort_values(COL_GNR_DT, ascending=False)

        if y < 40 * mm:
            c.showPage()
            y = height - 20 * mm
            c.setFont("Helvetica-Bold", 16)
            c.drawString(margin_x, y, f"ALL OTIF FAILED ORDERS — {year} (cont.)")
            y -= 12 * mm

        c.setFont("Helvetica-Bold", 12)
        header = f"{vendor_idx}. {vendor}   (Failures: {failures})   OTIF: {otif_pct:.1f}%   Contribution: {contrib_pct:.1f}%   Total Orders: {total_orders}"
        c.drawString(margin_x, y, header)
        y -= line_height

        c.setFont("Helvetica", 10)
        for _, orow in vendor_group.iterrows():
            ord_date = orow.get(COL_GNR_DT)
            date_str = "" if pd.isna(ord_date) else pd.to_datetime(ord_date).strftime("%d-%m-%Y")
            po_no = str(orow.get(COL_PO_NO, ""))
            line_text = f"    {date_str}    {po_no}"
            c.drawString(margin_x + 6 * mm, y, line_text)
            y -= (6 * mm)
            if y < 25 * mm:
                c.showPage()
                y = height - 20 * mm
                c.setFont("Helvetica", 10)
        y -= 4 * mm
        vendor_idx += 1

    c.save()
    buf.seek(0)
    with open(output_path, "wb") as f:
        f.write(buf.read())
    print(f"PDF written to: {output_path}")

# ---------- RUN: Upload file & process ----------
from google.colab import files
uploaded = files.upload()
if not uploaded:
    raise SystemExit("No file uploaded.")
# take the first uploaded file
file_name = list(uploaded.keys())[0]
print("Uploaded:", file_name)

# Read Excel with auto-detection
try:
    df_raw = read_excel_auto(file_name)
except Exception as e:
    # Provide the detailed exception so user can inspect
    raise RuntimeError(f"Error reading Excel file: {e}")

# Normalize column names to canonical names (accepts dotted or non-dotted variants)
df_raw = standardize_column_names(df_raw)
missing = [c for c in REQUIRED_COLS if c not in df_raw.columns]
if missing:
    raise ValueError(f"Input file missing required columns (after normalization): {missing}. Columns found: {list(df_raw.columns)}")

# Convert dtypes & drop nulls (will use canonical names)
df_raw = ensure_types_and_drop_nulls(df_raw)
if df_raw.empty:
    raise ValueError("No usable rows after type coercion / dropping nulls.")

# Merge Item Category mapping (if Item Category not present, merge from df1)
df = merge_item_category(df_raw)

# Lead-time rules: known defaults + user overrides (custom_lead_times)
lead_time_rules = DEFAULT_RULES.copy()
# apply overrides from custom_lead_times dict (if user configured)
if custom_lead_times:
    for k, v in custom_lead_times.items():
        lead_time_rules[str(k).strip().upper()] = int(v)

# Detect Mat Types present and auto-assign default lead times for unknown Mat Types
mat_types_in_data = set(df[COL_MAT_TYPE].dropna().astype(str).str.strip().str.upper().unique().tolist())
known_keys = set(lead_time_rules.keys()) | set(["PPM"])
unknowns = sorted(mat_types_in_data - known_keys)

if unknowns:
    print("Found Mat Types with no specified lead time. Assigning default lead time (days) to them:")
    for u in unknowns:
        print(f"  - {u} -> {DEFAULT_UNKNOWN_LEAD_TIME} days")
        lead_time_rules[u] = DEFAULT_UNKNOWN_LEAD_TIME

# Compute Lead Time for each row
df["Lead Time"] = df.apply(lambda row: compute_lead_time_for_row(row, lead_time_rules), axis=1)

# Final fallback — if any rows still have NaN lead time, fill with default and notify
if df["Lead Time"].isna().any():
    print("Warning: Some rows still lack Lead Time; filling with DEFAULT_UNKNOWN_LEAD_TIME.")
    print(df.loc[df["Lead Time"].isna(), [COL_MAT_TYPE]].drop_duplicates().head(10).to_string(index=False))
    df["Lead Time"] = df["Lead Time"].fillna(DEFAULT_UNKNOWN_LEAD_TIME)

# PO-level metrics
df_line, df_po = compute_po_level_metrics(df)

years = sorted(df_po["Year"].dropna().unique().astype(int).tolist())
if not years:
    raise ValueError("No valid years found in processed data.")
selected_year = years[-1]  # choose most recent by default
print("Selected year:", selected_year)

po_year = df_po[df_po["Year"] == selected_year].copy()

# Monthly summary
monthly = (
    po_year.groupby(["MonthNum", "Month"], as_index=False)
           .agg(
               Avg_OTIF=("OTIF", "mean"),
               Avg_OnTime=("OnTime", "mean"),
               Avg_InFull=("PO_Fulfilled", "mean"),
               Total_Orders=(COL_PO_NO, "count")
           )
           .sort_values("MonthNum")
)

# Vendor stats for the selected year
total_orders_year = po_year.shape[0]
vendor_stats = (
    po_year.groupby(COL_SUPPLIER, dropna=False)
           .agg(Total_Orders=(COL_PO_NO, "count"),
                OTIF_Failures=("OTIF", lambda x: int((x==0).sum())),
                OTIF_Success=("OTIF", lambda x: int((x==1).sum())))
           .reset_index()
)
vendor_stats[COL_SUPPLIER] = vendor_stats[COL_SUPPLIER].fillna("Unknown Supplier")
vendor_stats["Vendor_OTIF_pct"] = vendor_stats["OTIF_Success"] / vendor_stats["Total_Orders"] * 100
vendor_stats["Total_Contribution_pct"] = vendor_stats["Total_Orders"] / (total_orders_year if total_orders_year>0 else 1) * 100

# Top 10 vendors by failures
top10 = vendor_stats[vendor_stats["OTIF_Failures"]>0].sort_values("OTIF_Failures", ascending=False).head(10)
if top10.empty:
    print("No OTIF failures found in selected year.")
else:
    # Format and display top10
    display_top10 = top10[[COL_SUPPLIER, "OTIF_Failures", "Vendor_OTIF_pct", "Total_Contribution_pct", "Total_Orders"]].copy()
    display_top10["Vendor_OTIF_pct"] = display_top10["Vendor_OTIF_pct"].map(lambda v: f"{v:.1f}%")
    display_top10["Total_Contribution_pct"] = display_top10["Total_Contribution_pct"].map(lambda v: f"{v:.1f}%")
    print("\nTop vendors (failures, OTIF%, contribution%):")
    print(display_top10.to_string(index=False))

# Save CSVs
po_csv = f"po_level_{selected_year}.csv"
monthly_csv = f"monthly_otif_{selected_year}.csv"
po_year.to_csv(po_csv, index=False)
monthly.to_csv(monthly_csv, index=False)
print(f"Saved: {po_csv}, {monthly_csv}")

# Generate PDF for failures (grouped by supplier)
breaches = po_year[po_year["OTIF"] == 0].copy()
if breaches.shape[0] == 0:
    print("No OTIF breaches in selected year.")
else:
    out_pdf = f"OTIF_failed_orders_{selected_year}.pdf"
    generate_failed_orders_pdf_colab(breaches[[COL_SUPPLIER, COL_GNR_DT, COL_PO_NO]], vendor_stats, selected_year, out_pdf)
    # Offer file for download in Colab:
    from google.colab import files as gfiles
    gfiles.download(out_pdf)

print("Completed processing.")
