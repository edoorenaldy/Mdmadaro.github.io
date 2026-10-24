from __future__ import annotations

import csv
import io
import os
import time
from pathlib import Path
from typing import Dict, List, Tuple

import requests
from PIL import Image, ImageFile

ImageFile.LOAD_TRUNCATED_IMAGES = False
MAX_BYTES = 3000
OUT_DIR = Path("assets/images")
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; MaterialCatalogImageBot/1.0)"}

# Material Number | Material Description. All 148 rows were read from export_list_mn.xlsx.
ROWS_TEXT = r"""
56101536-0002|SOUND LEVEL METER ADJ TRIPOD_170CM_5KG
56101530-0017|BACKDROP CABINET_3000X500X2570MM
56101703-0046|RECEPTION DESK_3000X600X500MM
56101531-0012|SHOE RACK CABINET_40X26X89CM_4D_PLASTIC
56101519-0046|ROUND TABLE C/W 4-CHAIR_100X74CM_TOP MDF
56101601-0006|PATIO UMBRELLA W/ TRIPOD_DIA240XH230 CM
56101703-0043|OFFICE DESK_120X60X75CM
56101531-0011|SHOE RACK CABINET_116.5X35X115.3CM_4DOOR
56101530-0013|FILE CABINET_185X90X40CM_4-DOORS_METAL
56101530-0014|SLIDING FILE CABINET_185X90X40CM_METAL
56101519-0042|KITCHEN TABLE_2X0.7X0.85M_2 LAYER_S/S
56101519-0043|KITCHEN TABLE_1X0.7X0.85M_2 LAYER_S/S
56101508-0009|MATTRESS_200X90X15CM
56101520-0010|STEEL LOCKER_380X380X1830MM_2 DOOR
56101539-0005|BEDFRAME_120X200X36CM_MDF
56101504-0027|CHAIR_5.5X55X80CM_PP PLASTIC
56101519-0041|ROUND TABLE_DIA 60X75CM_MDF TOP
56101502-0020|PUFF SOFA_1200X500X450MM
56101502-0021|PUFF SOFA_650X650X420MM
56101504-0023|4-SEATER CHAIR_W/O BACKREST_IRON STEEL
56101504-0024|LOUNGE CHAIR_750X800X850MM
56101519-0039|COFFEE TABLE_DIA1000X380MM
56101519-0040|ROUND TABLE_800MM DIA_750MM HEIGHT
56101542-0003|FOLDING CHAIR C/W TABLE_WITH WHEEL
56101703-0038|OFFICE DESK WITHOUT SCREEN_2400X600MM
56101703-0039|OFFICE DESK WITHOUT SCREEN_3200X750MM
56101720-0009|OFFICE CHAIR_WIDE HEADREST_WITH WHEEL
56101502-0019|FABRIC SOFA_THREE-SEATER
56101519-0038|MEETING TABLE_200X120X75CM
56101531-0006|SHOE CABINET_190X40X135CM_SLIDING_ACP
56101531-0005|SHOE RACK_60X30X50_4 TIER_METAL
56101519-0037|TABLE_100X40X75 CM_3 TIER_WOOD
56101519-0035|ROUND TABLE_80X80X50CM_TOPGLASS 6MM
56101519-0036|ROUND TABLE_70X70X72CM
56101701-0004|CREDENZA_179.8X83.3X39CM
56101515-0025|SPRING BED /W BED FRAME_160 X 200CM
56101503-0001|COAT RACK_52CMX52CMX172CM_9 HOOK
56101504-0013|3-SEATER CHAIR WITHOUT BACKREST
56101703-0030|OFFICE DESK SET_3600X600MM
56101703-0031|OFFICE DESK SET_7200X1200MM
56101502-0010|SOFA BENCH_120X60MM
56101502-0011|SOFA FABRIC SET_1300X760X730MM
56101502-0012|SOFA FABRIC SET_1600X960X800MM
56101502-0013|SOFA FABRIC SET_TWO SEATER_CURVED
56101504-0010|2-SEATER CHAIR WITH TABLE
56101504-0011|1-SEATER CHAIR_WITH BACKREST_PLASTIC
56101504-0012|1-SEATER CHAIR_WITH_BACKREST_METAL
56101515-0019|SPRING BED_160X200 CM_CHIRO TYPE
56101519-0014|MEETING TABLE_1800X12000MM
56101519-0015|MEETING TABLE_2400X1000MM
56101519-0016|MEETING TABLE_2400X900MM
56101519-0017|MEETING TABLE_3600X1200MM
56101519-0018|ROUND TABLE_1000MM DIA
56101519-0019|ROUND TABLE_600MM DIA_700MM HEIGHT
56101519-0020|TABLE_600X1800MM
56101519-0021|TABLE_600X2400X750MM HEIGHT
56101519-0022|TABLE_700X700X750MM HEIGHT
56101519-0023|TABLE_700X700X750MM HEIGHT_METAL
56101520-0003|METAL LOCKER_18 DOOR_144X50X185 CM
56101520-0004|LOCKER_457X900X1830MM
56101701-0003|CREDENZA_385X60X76
56101703-0011|OFFICE DESK SET WITH SCREEN_1200X1000MM
56101703-0012|OFFICE DESK SET WITH SCREEN_1200X1200MM
56101703-0013|OFFICE DESK SET WITH SCREEN_1200X2400MM
56101703-0014|OFFICE DESK SET WITH SCREEN_1200X3600MM
56101703-0015|OFFICE DESK SET WITH SCREEN_1200X4800MM
56101703-0016|OFFICE DESK SET WITH SCREEN_1200X6000MM
56101703-0017|OFFICE DESK SET WITH SCREEN_1800X1200MM
56101703-0018|OFFICE DESK SET WITH SCREEN_600X4000MM
56101703-0019|OFFICE DESK SET WITH SCREEN_600X4800MM
56101703-0020|OFFICE DESK  W/ APART CREDENZA_70X150CM
56101703-0021|OFFICE DESK  W/ BACK CREDENZA_70X150CM
56101703-0022|OFFICE DESK  W/ JOIN CREDENZA_70X150CM
56101703-0023|OFFICE DESK W/ TOPBACK CREDENZA_70X150CM
56101703-0024|OFFICE DESK WITH SCREEN_750X3200MM
56101703-0025|OFFICE DESK WITHOUT SCREEN_1200X1200MM
56101703-0026|OFFICE DESK WITHOUT SCREEN_2000X600MM
56101703-0027|OFFICE DESK WITHOUT SCREEN_3600X600MM
56101703-0028|OFFICE DESK WITHOUT SCREEN_600X1200MM
56101703-0029|OFFICE DESK WITHOUT SCREEN_600X4000MM
56101515-0017|SPRING BED_180 X 75 CM
56101515-0018|SPRING BED_180 X 76 CM
56101502-0008|SOFA FABRIC_ONE SEATER
56101502-0009|SOFA FABRIC_TWO SEATER
56101504-0007|4-SEATER CHAIR_IRON STEEL
56101504-0008|4-SEATER CHAIR_WITH BACKREST_IRON STEEL
56101530-0006|FILE CABINET_390X800X2080_METAL
56101519-0006|ROUND TABLE_1000MM DIA_750MM HEIGHT
56101519-0007|MEETING TABLE_1200X2400MM
56101519-0009|MEETING TABLE_2000X800MM
56101703-0007|OFFICE DESK SET_600X1200MM
56101703-0008|OFFICE DESK WITHOUT SCREEN_600X1000MM
56101720-0003|OFFICE CHAIR HEADREST_WITH WHEEL
56101720-0005|OFFICE CHAIR_WITH WHEEL
56101720-0006|OFFICE CHAIR_WITHOUT WHEEL
56101504-0006|LEATHER CHAIR_58X54X106 CM_BLACK
56101702-0003|FILE CABINET SWING_4 DOORS_180X90X40 CM
56101504-0005|3-SEATER CHAIR_STAINLESS STEEL MATERIAL
56101703-0006|OFFICE DESK WITH CABINET_120 X 74 X 75CM
56101542-0001|FOLDING CHAIR_C/W TABLE
56101515-0015|SPRING BED_180 X 200CM
56101504-0004|CHAIR_420 X 505 X 868MM_CAESAR TYPE
56101537-0001|DRESSING TABLE_COMPLETE WITH CHAIR
56101515-0014|SPRING BED_100 X 200CM
56101539-0001|BUNK BED_90 X 200CM_170CM HEIGHT
56101515-0012|SPRING BED_70 X 180CM
56101519-0004|WORK TABLE_COMPLETE WITH CHAIR
56101532-0001|FURNITURE SET
56101504-0001|PLASTIC CHAIR
56101530-0001|FILING CABINET
56101502-0002|SOFA L SYNTH LEATHER P:65 L:60 T:77CM
56101502-0003|SOFA SYNTH LEATHER P:85 L:60 T :77CM
56101502-0005|SOFA SYNTH LEATHER P:180 L:60 T:77CM
56101515-0003|MATRESS_195 X 95 X 20CM
56101515-0004|SPRING BED_180 X 80CM
56101515-0005|SPRING BED_180 X 90CM
56101515-0006|SPRING BED_195 X 75CM
56101515-0007|SPRING BED_200 X 80CM
56101515-0008|SPRING BED_185 X 85CM
56101515-0009|SPRING BED_90 X 190CM
56101515-0010|SPRING BED_90 X 200CM
56101515-0011|SPRING BED_120 X 200CM
56101515-0013|MATRESS_200 X 120 X 20CM
56101504-0002|CHAIR D-320
56101504-0003|OFFICE CHAIR_BLACK
56101520-0001|LOCKER 6 DOOR 180X90X45CM
56101702-0002|FILE CABINET SLIDING 180X90X39CM
56101703-0002|DESK BEECH CT3C 240X120X75
56101703-0003|DESK HIGHPOINT ONE OD032 (120CM)-PD 133
56101515-0001|SINGLE BED_120 X 200CM
56101515-0002|BUNK BED_120 X 200CM
56101538-0001|FURNITURE - KITCHEN CABINET
56101604-0001|HAMMOCK
56101807-0001|BABY GIFT PACKAGE
56101502-0001|SOFA
56101508-0001|MATRAS
56101519-0002|MEETING TABLE
56101519-0003|WORK TABLE
56101543-0001|DINING TABLE
56101720-0004|CHAIR MEETING WITHOUT HANDSTAND
56101504-0000|CHAIR
56101515-0000|BED
56101520-0000|LOCKER
56101524-0000|IRONING BOARD
56101531-0000|SHOE RACK
56101542-0000|FOLDING CHAIR
56101702-0000|FILING CABINET
56101703-0000|DESK
"""

SOURCE_IMAGES: Dict[str, List[str]] = {
    "sound_meter": ["https://media.pt-cdn.de/production/public/media/d6/e9/bb/1616414910/8005_Stativ.jpg"],
    "reception_desk": ["https://cdn.shopify.com/s/files/1/0083/0436/1557/files/Untitled-1_30d16618-b9b9-40b8-b0a1-9512b06a4c02_600x600.png?v=1691714008"],
    "shoe_rack": [
        "https://down-tw.img.susercontent.com/file/sg-11134201-7rbkg-lq8917jbfs043d",
        "https://img.kilimall.com/c/obs/seller/3692/goods_image/251017111210_4c99afe563b4369df5d943ab38520ee9.png",
    ],
    "table_set": ["https://www.ewfdirect.com/cdn/shop/products/NDVA5-OAK-W_420513d0-8685-40d8-9472-a8a07b535c2d_1200x1200.jpg?v=1661979364"],
    "umbrella": [
        "https://static.homepro.co.th/seller/1251/hp-resize/product/resize_1000x1000_cover_img-1746514206252-0031DB.png",
        "https://www.ninesmw.com/cdn/shop/files/Nines_31960_648b479e-7004-4694-b233-47f2953966b6_1024x1024.jpg?v=1711993101",
    ],
    "desk": ["https://cdn11.bigcommerce.com/s-fkdbqpw9ac/images/stencil/1280x1280/products/10684/63858/desk-with-white-50mm-metal-legs__54872.1770306486.jpg?c=1"],
    "filing_cabinet": ["https://www.fubaofurniture.com/wp-content/uploads/2021/09/vertical-filing-cabinet-4-drawers-1.jpg"],
    "round_table": ["https://raw.githubusercontent.com/edoorenaldy/Mdmadaro.github.io/main/assets/images/meeting_table_reference.jpg"],
    "meeting_table": ["https://raw.githubusercontent.com/edoorenaldy/Mdmadaro.github.io/main/assets/images/meeting_table_reference.jpg"],
    "kitchen_table": ["https://ykp.sg/image/cache/catalog/2tier_table_logo-800x800.jpg"],
    "mattress": ["https://media.homeboxstores.com/i/homebox/167196533-167196533-HMBX15082024N_06-2100.jpg?%24prodimg-m-sqr-pdp-2x%24=&%24quality-standard%24=&fmt=auto&sm=c"],
    "locker": ["https://img.crazysales.com.au/products_pictures/378/378720_3416023_HD.jpg"],
    "bed": ["https://www.nitori-net.jp/ecstatic/image/product/2110600045108-0000005651994/2110600045108-000000565199401.jpg?ts=20260127105135762"],
    "bunk_bed": ["https://media.adeo.com/marketplace/MKP/91253740/d1c1ce9da36eb5d7aa5c76674de73c80.jpeg"],
    "plastic_chair": ["https://images.tcdn.com.br/img/img_prod/554605/cadeira_de_plastico_branca_empilhavel_bistro_antares_ate_154kg_93107_1_adc14cc2b0f5fea9fd4d6ac82ac17fc0.jpg"],
    "office_chair": ["https://bf1af2.akinoncloudcdn.com/products/2024/10/09/157138/716d15a3-d8e8-4dc3-8736-1325686f38ad.jpg"],
    "steel_chair": ["https://image.made-in-china.com/2f0j00LiUkYMEaqnqp/High-Quality-Stainless-Steel-Single-Seat-Chair-Hospital-Bank-Waiting-Area-Reception-Chairs.jpg"],
    "sofa": ["https://images.tcdn.com.br/img/img_prod/464234/sofa_3_lugares_para_sala_living_com_puff_280cm_morris_f04_linho_tramado_bege_lyam_decor_77751_1_0ae6dfa4677e5f4437a52a9fa95ff91c.jpg"],
    "lounge_chair": ["https://irrorwxhkpnjli5p-static.micyjz.com/cloud/ljBprKnplmSRilqipmrlkp/8036027.png"],
    "coat_rack": ["https://www.kapstokkendirect.nl/59884-large_default/spinder-design-charly-staande-kapstok-zwart.jpg"],
    "credenza": ["https://www.worthingtondirect.com/media/catalog/category/jsi-lok-storage-closed-without-top-hero.jpg"],
    "dressing_table": ["https://assets.wfcdn.com/im/67880125/resize-h800-w800%5Ecompr-r85/2551/255109002/Etchemin%2BDressing%2BTable%2Bwith%2BMirror-28824435.jpg"],
    "kitchen_cabinet": ["https://images.thdstatic.com/productImages/9aa6d37a-a7e8-42aa-8988-148ee7a8b676/svn/medium-oak-hampton-bay-assembled-kitchen-cabinets-kb33-mo-64_1000.jpg"],
    "hammock": ["https://www.2modern.com/cdn/shop/products/26302678_color-stripe-ocean-blue.jpg?v=1654145288"],
    "ironing_board": ["https://cdn.songmicswalldecor.com/cdn/10857/2025/10/31/main-SONGMICS-Ironing-Board-Foldable-Iron-Board-Height-Adjustable-up-to-35-Inches-with-4-Layered-Cover-and-Pad-Iron-Rest-Hook-Steel-XLegs-for-Home-Laundry--4039101-8847.jpg"],
    "baby_gift": ["https://tinywelcome.co.uk/cdn/shop/files/Neutral_Bew_Baby_Gift_Hamper_2.png?v=1769100633&width=1946"],
    "training_chair": ["https://icdn.tradew.com/file/202201/1575313/png/7729729.png?x-oss-process=image%2Fquality%2CQ_90"],
    # Material-specific image sources for the items that were reported as mismatched.
    "backdrop_cabinet": ["https://raw.githubusercontent.com/edoorenaldy/Mdmadaro.github.io/main/assets/images/backdrop_cabinet_reference.jpg"],
    "two_seater_chair_table": ["https://raw.githubusercontent.com/edoorenaldy/Mdmadaro.github.io/main/assets/images/two_seater_chair_table_reference.jpg", "https://ideacdn.net/idea/fm/97/myassets/products/603/961.jpg?revision=1734979179"],
    "office_desk_apart_credenza": ["https://www.soburo.com/19682-medium_default/bureau-individuel-sur-credence-so-vicar.jpg"],
    "office_desk_back_credenza": ["https://s7d9.scene7.com/is/image/NationalBusinessFurniture/BES_13496_6_BarkGraySlate?hei=750&wid=750"],
    "office_chair_without_wheel": ["https://ecdn6.globalso.com/upload/p/554/image_product/2023-12/656bc4d81da9142639.jpg"],
    "file_cabinet_sliding": ["https://img.gkbcdn.com/s3/p/2022-02-11/Office-Cabinet-with-Sliding-Door-Light-Grey-90x40x180-cm-Steel-496409-8.jpg"],
    "meeting_chair_without_armrest": ["https://raw.githubusercontent.com/edoorenaldy/Mdmadaro.github.io/main/assets/images/meeting_chair_without_armrest_reference.jpg", "https://www.ypo.co.uk/cdn-cgi/image/format%3Dauto%2Cquality%3D75%2Cwidth%3D568%2Cheight%3D568/https%3A/dl.ypo.co.uk/website_product_images/D5/d55790_01_2023.jpg"],
    "mattress_200_90_15": ["https://cdn.webshopapp.com/shops/307649/files/445994965/viking-choice-schuim-matras-h3-90x200cm-dikte-15cm.jpg"],
    "bed_frame_120_200": ["https://www.sconto.cz/content/images/product/variantRetina/arizona-iii_390112.jpg", "https://gomagcdn.ro/domains/bortisimpex.ro/files/product/original/pat-120x200-stejar-artisan-airon-320923.png"],
    "round_table_60_70": ["https://homemart.co.nz/cdn/shop/products/124bf6d6ae937d15bd4a95fc10cf09f3.jpg?v=1681430578&width=1090"],
    "spring_mattress_75": ["https://www.ikea.com.tw/dairyfarm/tw/images/508/1150829_PE884908_S5.jpg", "https://www.wilko.com/assets/bWFzdGVyfGltYWdlc3wyMDc5Mzd8aW1hZ2UvanBlZ3xoYWEvaDMyLzk4MjY1NTYxODI1NTgvMDcyODg3MF82LmpwZ3xiNjg0NTI0ODJjOTI1NjI4YzE3Yjg2MzIwNTUxNDZkNjhlZTZmODBmMzUzZjIwZmQzZTVlYTExYWU4Y2I4ZWU4/0728870-6.jpg", "https://image.rakuten.co.jp/nitori/cabinet/20101/2010118_page1_014.jpg"],
    "mattress_195_95_20": ["https://kurashiec202009.itembox.design/product/093/000000009390/000000009390-13-l.jpg?t=20220930162538"],
    "single_bed_120_200": ["https://cdn.dedeman.ro/media/catalog/product/8/1/8128907_1.jpg?canvas=%3A&fit=bounds&height=&optimize=low&width="],
}

def categorize(material: str, desc: str) -> str:
    d = desc.upper()
    # Exact Material Number overrides prevent similar furniture categories from sharing the wrong photo.
    material_specific = {
        "56101530-0017": "backdrop_cabinet",
        "56101504-0010": "two_seater_chair_table",
        "56101703-0020": "office_desk_apart_credenza",
        "56101703-0021": "office_desk_back_credenza",
        "56101720-0006": "office_chair_without_wheel",
        "56101702-0002": "file_cabinet_sliding",
        "56101720-0004": "meeting_chair_without_armrest",
        "56101508-0009": "mattress_200_90_15",
        "56101539-0005": "bed_frame_120_200",
        "56101519-0019": "round_table_60_70",
        "56101515-0017": "spring_mattress_75",
        "56101515-0018": "spring_mattress_75",
        "56101515-0003": "mattress_195_95_20",
        "56101515-0001": "single_bed_120_200",
    }
    if material in material_specific:
        return material_specific[material]
    if "SOUND LEVEL METER" in d:
        return "sound_meter"
    if "BABY GIFT" in d:
        return "baby_gift"
    if "IRONING BOARD" in d:
        return "ironing_board"
    if "HAMMOCK" in d:
        return "hammock"
    if "PATIO UMBRELLA" in d:
        return "umbrella"
    if "RECEPTION DESK" in d:
        return "reception_desk"
    if "SHOE RACK" in d or "SHOE CABINET" in d:
        return "shoe_rack"
    if "DRESSING TABLE" in d:
        return "dressing_table"
    if "KITCHEN CABINET" in d:
        return "kitchen_cabinet"
    if "COAT RACK" in d:
        return "coat_rack"
    if "CREDENZA" in d or "BACKDROP CABINET" in d:
        return "credenza"
    if "LOCKER" in d:
        return "locker"
    if "FILING CABINET" in d or "FILE CABINET" in d:
        return "filing_cabinet"
    if "FURNITURE SET" in d:
        return "table_set"
    if "TABLE C/W 4-CHAIR" in d:
        return "table_set"
    if "KITCHEN TABLE" in d:
        return "kitchen_table"
    if "BUNK BED" in d:
        return "bunk_bed"
    if "MATTRESS" in d or "MATRESS" in d or "MATRAS" in d:
        return "mattress"
    if "BED" in d or material.startswith("56101515-") or material.startswith("56101539-"):
        return "bed"
    if "FOLDING CHAIR" in d:
        return "training_chair"
    if "OFFICE CHAIR" in d or material.startswith("56101720-"):
        return "office_chair"
    if "SOFA" in d or "PUFF" in d or "SOFA BENCH" in d:
        return "sofa"
    if "LOUNGE CHAIR" in d or "LEATHER CHAIR" in d:
        return "lounge_chair"
    if "STAINLESS STEEL" in d or "IRON STEEL" in d or "WITH_BACKREST_METAL" in d:
        return "steel_chair"
    if "PLASTIC CHAIR" in d or ("CHAIR" in d and "PLASTIC" in d):
        return "plastic_chair"
    if material.startswith("56101504-"):
        return "plastic_chair"
    if material.startswith("56101502-"):
        return "sofa"
    if material.startswith("56101701-"):
        return "credenza"
    if material.startswith("56101702-") or material.startswith("56101530-"):
        return "filing_cabinet"
    if material.startswith("56101703-"):
        return "desk"
    if "ROUND TABLE" in d or "COFFEE TABLE" in d:
        return "round_table"
    if "TABLE C/W 4-CHAIR" in d:
        return "table_set"
    if "TABLE" in d and "DINING" in d:
        return "table_set"
    if material.startswith("56101519-") or "TABLE" in d or "DESK" in d:
        return "meeting_table" if "MEETING" in d else "desk"
    return "desk"

FALLBACK_CATEGORIES: Dict[str, List[str]] = {
    "reception_desk": ["desk", "credenza"],
    "shoe_rack": ["locker", "filing_cabinet"],
    "table_set": ["round_table", "meeting_table"],
    "desk": ["reception_desk", "credenza"],
    "filing_cabinet": ["credenza", "locker"],
    "round_table": ["meeting_table", "table_set"],
    "meeting_table": ["desk", "round_table"],
    "kitchen_table": ["meeting_table", "desk"],
    "mattress": ["bed"],
    "locker": ["filing_cabinet", "credenza"],
    "bed": ["mattress"],
    "bunk_bed": ["bed", "mattress"],
    "plastic_chair": ["steel_chair", "training_chair"],
    "office_chair": ["training_chair", "plastic_chair"],
    "steel_chair": ["plastic_chair", "training_chair"],
    "sofa": ["lounge_chair"],
    "lounge_chair": ["sofa"],
    "credenza": ["filing_cabinet", "kitchen_cabinet"],
    "dressing_table": ["desk", "credenza"],
    "kitchen_cabinet": ["credenza", "filing_cabinet"],
    "training_chair": ["office_chair", "plastic_chair"],
    "ironing_board": ["desk"],
    "backdrop_cabinet": ["credenza", "filing_cabinet"],
    "two_seater_chair_table": ["steel_chair", "plastic_chair"],
    "office_desk_apart_credenza": ["desk", "credenza"],
    "office_desk_back_credenza": ["desk", "credenza"],
    "office_chair_without_wheel": ["steel_chair", "training_chair"],
    "file_cabinet_sliding": ["filing_cabinet", "locker"],
    "meeting_chair_without_armrest": ["training_chair", "steel_chair"],
    "mattress_200_90_15": ["mattress"],
    "bed_frame_120_200": ["bed"],
    "round_table_60_70": ["round_table"],
    # No fallback for spring mattresses: avoid substituting a bed-with-frame photo.
    "mattress_195_95_20": ["mattress"],
    "single_bed_120_200": ["bed"],
}

def fetch_image(session: requests.Session, category: str) -> Tuple[Image.Image, str]:
    errors = []
    source_categories = [category] + FALLBACK_CATEGORIES.get(category, [])
    for source_category in source_categories:
        for url in SOURCE_IMAGES.get(source_category, []):
            try:
                response = session.get(url, timeout=(4, 8), headers=HEADERS)
                response.raise_for_status()
                if "image" not in response.headers.get("Content-Type", "").lower() and not response.content.startswith((b"\xff\xd8", b"\x89PNG", b"RIFF")):
                    raise ValueError("Response is not an image")
                source = Image.open(io.BytesIO(response.content))
                source.load()
                if source.mode in ("RGBA", "LA"):
                    background = Image.new("RGB", source.size, "white")
                    background.paste(source, mask=source.getchannel("A"))
                    source = background
                else:
                    source = source.convert("RGB")

                # Remove measurement callouts/marketing copy from representative photos where possible.
                w, h = source.size
                if category == "bed_frame_120_200":
                    source = source.crop((int(w * 0.10), int(h * 0.24), int(w * 0.90), int(h * 0.68)))
                elif category == "mattress_195_95_20":
                    source = source.crop((int(w * 0.02), int(h * 0.20), int(w * 0.89), int(h * 0.60)))
                elif category == "spring_mattress_75":
                    if "wilko.com" in url:
                        source = source.crop((int(w * 0.02), int(h * 0.20), int(w * 0.95), int(h * 0.67)))
                    elif "rakuten.co.jp" in url:
                        source = source.crop((int(w * 0.22), int(h * 0.50), int(w * 0.99), int(h * 0.93)))

                source.thumbnail((120, 120), Image.Resampling.LANCZOS)
                return source, url
            except Exception as exc:
                errors.append(f"{url}: {exc}")
    raise RuntimeError(f"Could not retrieve image for {category}: " + " | ".join(errors))
def make_jpeg(image: Image.Image) -> bytes:
    for dimension in (120, 100, 90, 80, 72, 64):
        resized = image.copy()
        resized.thumbnail((dimension, dimension), Image.Resampling.LANCZOS)
        for quality in (78, 68, 58, 48, 38, 28, 20):
            buff = io.BytesIO()
            resized.save(buff, format="JPEG", quality=quality, optimize=True, progressive=False, subsampling=2)
            data = buff.getvalue()
            if len(data) <= MAX_BYTES:
                return data
    raise RuntimeError("Unable to compress image below 3000 bytes")

def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    material_rows = []
    for line in ROWS_TEXT.strip().splitlines():
        if "|" not in line:
            continue
        material, description = line.split("|", 1)
        material_rows.append((material.strip(), description.strip()))
    if len(material_rows) != 148:
        raise RuntimeError(f"Expected 148 materials, got {len(material_rows)}")

    session = requests.Session()
    category_cache: Dict[str, Tuple[bytes, str]] = {}
    results = []
    failures = []
    for idx, (material, description) in enumerate(material_rows, start=1):
        category = categorize(material, description)
        try:
            if category not in category_cache:
                image, source_url = fetch_image(session, category)
                category_cache[category] = (make_jpeg(image), source_url)
            data, source_url = category_cache[category]
            out = OUT_DIR / f"{material}.jpg"
            out.write_bytes(data)
            # Verify each output can be opened as a valid JPEG.
            with Image.open(out) as check:
                check.verify()
            results.append((material, description, category, out.as_posix(), source_url, len(data), "OK"))
            print(f"[{idx}/148] OK {material} ({category}, {len(data)} bytes)")
        except Exception as exc:
            failures.append((material, description, category, str(exc)))
            print(f"[{idx}/148] ERROR {material}: {exc}")

    csv_path = OUT_DIR / "MATERIAL_IMAGE_SOURCES.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["Material Number", "Material Description", "Image Category", "Image File", "Image Source URL", "File Size (bytes)", "Status"])
        writer.writerows(results)
        for material, description, category, error in failures:
            writer.writerow([material, description, category, f"assets/images/{material}.jpg", "", "", "FAILED: " + error[:250]])

    if failures:
        print(f"Completed with {len(failures)} failures. See {csv_path}.")
        # Workflow should still commit successful images and source list.
    print(f"Successful images: {len(results)}; failed: {len(failures)}; unique categories: {len(category_cache)}")
    if not results:
        raise RuntimeError("No images were downloaded successfully")

if __name__ == "__main__":
    main()
