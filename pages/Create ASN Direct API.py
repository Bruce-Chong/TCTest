from databricks import sql
import os
from dotenv import load_dotenv
import streamlit as st
import pandas as pd
import os
from pathlib import Path
import openpyxl
from datetime import datetime, timedelta
import requests
import json
import re
import uuid
from decimal import Decimal


# Load environment variables from .env file
load_dotenv()

# Function to connect to Databricks and fetch data
import databricks.sql as sql
import streamlit as st

db_token = os.getenv("DATABRICKS_TOKEN")
db_hostname = os.getenv("DATABRICKS_SERVER_HOSTNAME")
db_http = os.getenv("DATABRICKS_HTTP_PATH")

# set N API key here
client_id = os.getenv("CLIENT_ID")
client_secret = os.getenv("CLIENT_SECRET")
grant_type = "client_credentials"
tokendata = {
    "grant_type": grant_type,
    "client_id": client_id,
    "client_secret": client_secret
}
nike_auth_url = os.getenv("NIKE_AUTH_URL")
FHQ_URL = os.getenv("FHQ_URL")

def get_auth_header(auth_url):
    auth_response = requests.post(nike_auth_url, data=tokendata)
    token = json.loads(auth_response.text)['access_token']
    contentype = "'Content-Type': 'application/json'"
    headers = { 'Authorization' : 'Bearer ' +token , 'Content-type': contentype}
    return headers

def fetch_data(l_po, zpo_itm):
    try:
        if not l_po:
            st.warning("No PO numbers provided.")
            return None

        # Prepare PO header and item number filters
        po_list_str = ",".join(f"'{po}'" for po in l_po)
        item_list_str = ",".join(f"'{item}'" for item in zpo_itm) if zpo_itm else None

        # Build the query with both filters
        query = f"""
            SELECT DISTINCT 
                a.po_header_nbr, 
                a.po_item_nbr, 
                b.plant_cd,
                b.po_shipping_instruction_cd,
                b.product_cd,
                a.size_cd,
                a.po_on_order_qty,
                a.order_qty_uom,
                b.request_tracking_nbr
            FROM development.perf_purchase_order.curated_po_item_size_schedule_line_v a
            JOIN development.perf_purchase_order.curated_po_item_v b
                ON a.po_header_nbr = b.po_header_nbr
            WHERE a.po_header_nbr IN ({po_list_str})
              AND a.po_on_order_qty <> 0
        """

        # Add item number filter if provided
        if item_list_str:
            query += f" AND a.po_item_nbr IN ({item_list_str})"

        # Connect to Databricks
        connection = sql.connect(
            server_hostname=db_hostname,
            http_path=db_http,
            access_token=db_token
        )

        cursor = connection.cursor()
        cursor.execute(query)
        data = cursor.fetchall()

        cursor.close()
        connection.close()

        return data

    except Exception as e:
        st.error(f"An error occurred: {e}")
        return None



st.set_page_config(page_title='Create ASN')
st.title('Create ASN')
st.subheader('Choose an action')

# Path Setting:
try:
    current_path = Path(__file__).parent.absolute()
except:
    current_path = Path.cwd()

filepath = os.path.join(current_path, 'TestData.xlsx')

# Load Excel sheets
wb = openpyxl.load_workbook(filepath)

shipcode_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='ship_code')
poo_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='POO')
pod_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='POD')
plant_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='plant_add')
fac_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='factory')

# Set values for account for easy search
l_fac = fac_df['vendor'].tolist()
options = fac_df['vendor'].tolist()
dic = dict(zip(options, l_fac))

# Set API key and token (commented out for now)
client_id = "nike.sapcp.apim"
client_secret = "secret here"
grant_type = "client_credentials"
tokendata = {
    "grant_type": grant_type,
    "client_id": client_id,
    "client_secret": client_secret
}

# auth_response = requests.post(nike_auth_url, data=tokendata)
# token = json.loads(auth_response.text)['access_token']
# headers = { 'Authorization' : 'Bearer ' + token , 'Content-type': 'application/json' }

# Set today's date
today_dt = datetime.now()
tmr_dt = today_dt + timedelta(days=1)
today_dt = today_dt.strftime("%Y-%m-%d")
today_dts = str(today_dt) + "T01:00:00Z"
tmr_dt = tmr_dt.strftime("%Y-%m-%d")
st.session_state.today_dt = str(today_dt)

# The rest of your code (functions like is_number, update_field, post_api, and the form) remains unchanged
###############set all session state variables###############
st.session_state.today_dt = str(today_dt)


def is_number(s):
    if (s is None):
        return None
    try:
        float(s)
        return str(int(s))
    except ValueError:
        return str(s)

def update_field(zpo, zpo_itm, zfty):
    empty_del = {"deliveryItems": [{}]}
    total_ship = {"deliveries": []}
    total_gh = {"goodsHolders": []}

    shipped_from = is_number(fac_df.loc[fac_df['vendor'] == zfty, 'MCO'].iloc[0])

    deliveryno = 1
    ghno = 10009715482305423101

    # Define the expected column names in the correct order
    ekpo_columns = [
        'EBELN',  # 0
        'EBELP',  # 1
        'WERKS',  # 2
        'EVERS',  # 3
        'MATNR',  # 4
        'TXZ01',  # 5
        'QTY',
        'UOM',
        'DTR'
        # Add more column names here if your data has more fields
        # For example:
        # 'QUANTITY', 'UOM', 'NODECODE', 'GROSSWEIGHT', ...
    ]

    ekpo_data = fetch_data(zpo, zpo_itm)

    # Create the DataFrame with column names
    ekpo_df = pd.DataFrame(ekpo_data, columns=ekpo_columns)

    # Group by EBELN and EBELP
    grouped = ekpo_df.groupby(['EBELN', 'EBELP'])

    for (EBELN, EBELP), group in grouped:
        #MATNR = MATNR
        ITMNO = str(EBELP).zfill(5)

        deliveryItems = []
        goodsHolders = []

        deliveryitmno = 1

        for _, row in group.iterrows():
            MATNR = row['MATNR'].replace(" ", "")
            size = row['TXZ01'].replace(" ", "")
            qty = int(Decimal(str(row['QTY']).strip("Decimal('')")))
            uom = row['UOM'].replace(" ", "")  # Adjust if column index has changed
            dtr = row['DTR'].replace(" ", "")
            plant = row['WERKS'].replace(" ", "")

            deliveryItems.append({
                "deliveryNoteItemNumber": str(deliveryitmno),
                "productCode": MATNR,
                "sizeCode": size,
                "qualityCode": "01",
                "inventorySegmentationCode": "000",
                "deliveryQuantity": qty,
                "uOM": uom,
                "originCountryCode": shipped_from
            })

            goodsHolders.append({
                "goodsHolderTypeCode": "CARTON",
                "goodsHolderTypeSizeCode": "B10",
                "goodsHolderNumber": str(ghno),
                "goodsHolderItems": [
                    {
                        "productCode": MATNR,
                        "sizeCode": size,
                        "packedQuantity": qty,
                        "deliveryNoteNumber": str(deliveryno),
                        "deliveryNoteItemNumber": str(deliveryitmno)
                    }
                ]
            })

            deliveryitmno += 1
            ghno += 1

        delivery = {
            "deliveryNoteNumber": str(deliveryno),
            "receiptId": st.session_state.RID,
            "plannedGoodsReceiptDate": tmr_dt,
            "assignedNodeCode": plant,
            "deliveryReferenceAttributes": [
                {"referenceTypeCode": "ORIGINAL_DELIVERY_NUMBER", "referenceText": st.session_state.TCCI},
                {"referenceTypeCode": "INVOICE_DATE", "referenceText": today_dt},
                {"referenceTypeCode": "INVOICE_NUMBER", "referenceText": st.session_state.TCCI},
                {"referenceTypeCode": "PURCHASEORDER_NUMBER", "referenceText": str(EBELN)},
                {"referenceTypeCode": "PURCHASEORDER_ITEM_NUMBER", "referenceText": ITMNO}
            ],
            "deliveryMeasures": {
                "grossWeight": {
                    "measureValue": "1",
                    "measureUOM": "KG"
                },
                "grossVolume": {
                    "measureValue": "4.149",
                    "measureUOM": "M3"
                }
            },
            "deliveryItems": deliveryItems
        }

        total_ship["deliveries"].append(delivery)
        total_gh["goodsHolders"].extend(goodsHolders)

        deliveryno += 1

    post_api(total_ship, zpo, zfty, total_gh, ekpo_df)
    return None


def post_api(ttl_ship, zpo, zfty, ttl_gh, ekpo_df):
    # url and data definition
    base_url = "https://nikecfqaapiportal.prod.apimanagement.us20.hana.ondemand.com:443/delivery/v1/TCOutboundDelivery"
    st.dataframe(ekpo_df)
    for x in zpo:
        # if x != '[' and x != ']' and x != "'":
        st.write(x)
        zplant = is_number(ekpo_df.loc[ekpo_df['EBELN'] == str(x), 'WERKS'].iloc[0])
        VehicleTypeCode = ekpo_df.loc[ekpo_df['EBELN'] == str(x), 'EVERS'].iloc[0]

        filtered_df = pod_df.loc[pod_df['WERKS'] == int(zplant), VehicleTypeCode]
        if not filtered_df.empty:
            zpod = is_number(filtered_df.iloc[0])
        else:
            st.error(f"No matching data found for Port of Destination: {zplant}")
            zpod = st.text_input('Enter value for port of destination, if VL, e.g USMEM, or MEM for none VL')

        zshipcode = is_number(shipcode_df.loc[shipcode_df['ShipMode'] == VehicleTypeCode, 'ShipCode'].iloc[0])
        zpoo = is_number(poo_df.loc[poo_df['Factory'] == zfty, VehicleTypeCode].iloc[0])
        vendorCode = is_number(shipcode_df.loc[shipcode_df['ShipMode'] == VehicleTypeCode, 'LSPCode'].iloc[0])

    source = {
        "event": {
            "id": st.session_state.UUID,
            "timestamp": today_dts,
            "actionCode": "CREATE",
            "sourceSystemName": "SAP-AFS"
        },
        "shipment": {
            "shipmentTypeCode": "INBOUND",
            "bolNumber": "FHR1a",
            "proNumber": "FHR1a",
            "transportVehicleTypeCode": VehicleTypeCode,
            "legIndicatorCode": "4",
            "shipmentVoyageNumber": "Voyage",
            "totalNumberOfContainerCount": 1,
            "goodsHolderCount": 52,
            "costCalculationStatusCode": "B",
            "multiLegBOLIndicator": "true",
            "shipmentDates": {
                "actualShippedTimestamp": today_dts,
                "plannedDischargeDate": st.session_state.PDD,
                "estimatedDeliveryTimestamp": st.session_state.EDT
            },
            "shipmentDischargeAddress": {
                "shipmentDischargeCode": zpod,
                "shipmentDischargeTypeCode": zshipcode
            },
            "shipmentOriginAddress": {
                "shipmentOriginCode": zpoo,
                "shipmentOriginTypeCode": zshipcode
            },
            "shipmentDestinationAddress": {
                "shipmentDestinationCode": zplant,
                "shipmentDestinationTypeCode": "DC"
            },
            "shipmentVendors": [
                {
                    "vendorTypeCode": "FORWARDER",
                    "vendorCode": vendorCode
                },
                {
                    "vendorTypeCode": "CARRIER",
                    "vendorCode": vendorCode
                }
            ],
            "shipmentMeasures": {
                "grossWeight": {
                    "measureValue": "520.38",
                    "measureUOM": "KG"
                },
                "grossVolume": {
                    "measureValue": "520.83",
                    "measureUOM": "M3"
                }
            },
            "deliveries": [{}],
            "goodsHolders": [{}],
            "transportHolders": [
                {
                    "transportHolderNumber": "CONTAINER2_EC1",
                    "transportHolderTypeSizeCode": "9002",
                    "transportHolderSequenceNumber": "1",
                    "transportVehicleCraftCode": "Vessel name3",
                    "sealNumber": [
                        "Seal6"
                    ],
                    "multipleBillOfLadingIndicator": "true",
                    "loadLocationTypeCode": "C",
                    "transportHolderTypeCode": "CONTAINER",
                    "transportHolderMeasures": {
                        "grossWeight": {
                            "measureValue": "12.38",
                            "measureUOM": "KG"
                        },
                        "grossVolume": {
                            "measureValue": "10.832",
                            "measureUOM": "M3"
                        }
                    }
                }
            ]
        }
    }
    source["shipment"]["deliveries"] = ttl_ship["deliveries"]
    source["shipment"]["goodsHolders"] = ttl_gh["goodsHolders"]
    st.write(source)
    return source
###################create GUI##################

with st.form("Creat ASN"):
    st.write("No need for PO nor TestData.xls")
    #in_UUID = st.text_input("UUID")

    # Streamlit input field with default value

    auth_url = st.text_input("FHQ API URL", value=FHQ_URL)

    in_RID = st.text_input("Receipt ID")
    in_PO = st.text_input("CRPO number. Use , for more than one(no space).  IMPT! Do not put more than one PO if you enter PO item number")
    in_POitm = st.text_input("PO item number like 100 or 200 with no zeros in front, leave blank for all item. Use , for more than one(no space)")
    in_TCCI = st.text_input("TCCI number")
    in_PDD = st.text_input("Planned Discharge Date like 2023-08-23")
    in_EDT = st.text_input("estimated Delivery Time stamp like 2023-07-31")
    in_FTY = st.selectbox("Factory", options, format_func=lambda x: dic[x])

    # Checkbox as a parameter
    post_flag = st.checkbox("Direct Post to API")

    # Every form must have a submit button.
    submitted = st.form_submit_button("Create ASN from PO")

    if submitted:
        l_po = in_PO.split(",")

        # Split the input by commas and strip whitespace
        po_items = [item.strip() for item in in_POitm.split(',') if item.strip()]

        # Generate a random UUID (version 4)
        st.session_state.UUID = str(uuid.uuid4())
        st.session_state.RID = in_RID
        st.session_state.TCCI = in_TCCI
        st.session_state.PDD = in_PDD
        st.session_state.EDT = str(in_EDT) + "T01:00:00Z"
        try:
            payload = update_field(l_po, po_items, str(in_FTY))
            if post_flag:
                headers = get_auth_header(auth_url)
                r = requests.request("POST", auth_url, headers=headers, data=payload)
                print(r.text)

        except Exception as E:
            st.write(E)






