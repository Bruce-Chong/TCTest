
import streamlit as st
import pandas as pd
import os
from pathlib import Path
import openpyxl
from datetime import datetime, timedelta
import requests
import json
import re

st.set_page_config(page_title='Create ASN')
st.title('Create ASN')
st.subheader('Choose an action')


##################retrieve info from excel ***************************

# Path Setting:
try:
    current_path = Path(__file__).parent.absolute()  # Get the file path for this py file
except:
    current_path = (Path.cwd())

#filepath = os.path.join(current_path, 'TestData.xlsx')
if 'file' not in st.session_state:
    st.write("Please set file path at Menu first!")
else:
    filepath = st.session_state.file

wb = openpyxl.load_workbook(filepath)
ekpo_sheet = wb['EKPO']
ekpo_sheet.delete_rows(1)
#put worksheet into dataframe
shipcode_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='ship_code')
poo_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='POO')
pod_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='POD')
plant_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='plant_add')
fac_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='factory')
ekpo_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='EKPO')

# set values for account for easy search
l_fac = fac_df['vendor'].tolist()
options = fac_df['vendor'].tolist()
dic = dict(zip(options, l_fac))

# set API key here
client_id = "nike.sapcp.apim"
client_secret = "secret here"
grant_type = "client_credentials"
tokendata = {
    "grant_type": grant_type,
    "client_id": client_id,
    "client_secret": client_secret
}

#get token
nike_auth_url = "url here"
auth_response = requests.post(nike_auth_url, data=tokendata)
token = json.loads(auth_response.text)['access_token']
#print (token)
contentype = "'Content-Type': 'application/json'"
headers = { 'Authorization' : 'Bearer ' +token , 'Content-type': contentype}

######Set today's date###############
today_dt = datetime.now()
tmr_dt = today_dt + timedelta(days=1)
today_dt = today_dt.strftime("%Y-%m-%d")
today_dts = str(today_dt) + "T01:00:00Z"
tmr_dt = tmr_dt.strftime("%Y-%m-%d")

###############set all session state variables###############
st.session_state.today_dt = str(today_dt)


def is_number(s):
  if(s is None):
    return None
  try:
    float(s)
    return str(int(s))
  except ValueError:
    return str(s)


def update_field(zpo, zfty):

    empty_del = {"deliveryItems": [{}]}
    total_del = {"deliveryItems": [{}]}
    empty_ship = {"deliveries": [{}]}
    total_ship = {"deliveries": [{}]}
    empty_gh = {"goodsHolders": [{}]}
    total_gh = {"goodsHolders": [{}]}

    shipped_from = is_number(fac_df.loc[fac_df['vendor'] == zfty, 'MCO'].iloc[0])

    masteritm = "0"
    lineno = 0
    deliveryno = 1
    ghno = 10009715482305423101
    for row in ekpo_sheet.iter_rows():
        EBELN = is_number(row[0].value)
        EBELP = is_number(row[1].value)
        TXZ01 = is_number(row[6].value)
        size = TXZ01.split(",",1)
        if str(EBELN) in zpo:
            if EBELP[-2:] == "00":
                MATNR = is_number(row[7].value)
                ITMNO = EBELP.zfill(5)
                jship = {
                    "deliveries": [
                    {
                        "deliveryNoteNumber": str(deliveryno),
                        "receiptId": st.session_state.RID,
                        "plannedGoodsReceiptDate": tmr_dt,
                        "assignedNodeCode": is_number(row[10].value),
                        "deliveryReferenceAttributes": [
                            {
                                "referenceTypeCode": "ORIGINAL_DELIVERY_NUMBER",
                                "referenceText": st.session_state.TCCI
                            },
                            {
                                "referenceTypeCode": "INVOICE_DATE",
                                "referenceText": today_dt
                            },
                            {
                                "referenceTypeCode": "INVOICE_NUMBER",
                                "referenceText": st.session_state.TCCI
                            },
                            {
                                "referenceTypeCode": "PURCHASEORDER_NUMBER",
                                "referenceText": str(EBELN)
                            },
                            {
                                "referenceTypeCode": "PURCHASEORDER_ITEM_NUMBER",
                                "referenceText": ITMNO
                            }
                        ],
                        "deliveryMeasures": {
                            "grossWeight": {
                                "measureValue": is_number(row[26].value),
                                "measureUOM": "KG"
                            },
                            "grossVolume": {
                                "measureValue": "4.149",
                                "measureUOM": "M3"
                            }
                        },
                        "deliveryItems": [{}]
                    }
                    ],
                }
                # only do this if there is more than one master item
                if masteritm != "0":
                    total_ship["deliveries"][lineno]["deliveryItems"] = total_del["deliveryItems"]
                    total_ship["deliveries"] += jship["deliveries"]
                    total_del = empty_del
                    lineno = lineno + 1

                else:
                    total_ship["deliveries"] = jship["deliveries"]

                masteritm = EBELP
                deliveryno = deliveryno + 1
                deliveryitmno = 1
                
            else:
                deliveryItems = {
                    "deliveryItems": [
                        {
                            "deliveryNoteItemNumber": str(deliveryitmno),
                            "productCode": MATNR,
                            "sizeCode": size[1].replace(" ", ""),
                            "qualityCode": "01",
                            "inventorySegmentationCode": "000",
                            "deliveryQuantity": is_number(row[17].value),
                            "uOM": is_number(row[18].value),
                            "originCountryCode": shipped_from
                        }
                    ]
                }

                if total_del == empty_del:
                    total_del["deliveryItems"] = deliveryItems["deliveryItems"]
                else:
                    total_del["deliveryItems"] += deliveryItems["deliveryItems"]

                goodsHolders = {
                    "goodsHolders": [
                        {
                            "goodsHolderTypeCode": "CARTON",
                            "goodsHolderTypeSizeCode": "B10",
                            "goodsHolderNumber": str(ghno),
                            "goodsHolderItems": [
                                {
                                    "productCode": MATNR,
                                    "sizeCode": size[1].replace(" ", ""),
                                    "packedQuantity": is_number(row[17].value),
                                    "deliveryNoteNumber": str(deliveryno-1),
                                    "deliveryNoteItemNumber":str(deliveryitmno)
                                }
                            ]
                        }
                    ]
                }
                deliveryitmno = deliveryitmno + 1
                ghno = ghno + 1
                if total_gh == empty_gh:
                    total_gh["goodsHolders"] = goodsHolders["goodsHolders"]
                else:
                    total_gh["goodsHolders"] += goodsHolders["goodsHolders"]

    total_ship["deliveries"][lineno]["deliveryItems"] = total_del["deliveryItems"]
    #st.write(total_po)

    post_api(total_ship, zpo, zfty, total_gh)
    return None

def post_api(ttl_ship,zpo, zfty,ttl_gh):
    # url and data definition
    base_url = "https://nikecfqaapiportal.prod.apimanagement.us20.hana.ondemand.com:443/delivery/v1/TCOutboundDelivery"

    for x in zpo:
        #if x != '[' and x != ']' and x != "'":
        st.write(x)
        zplant = is_number(ekpo_df.loc[ekpo_df['EBELN'] == int(x), 'WERKS'].iloc[0])
        VehicleTypeCode = ekpo_df.loc[ekpo_df['EBELN'] == int(x), 'EVERS'].iloc[0]
        zpod = is_number(pod_df.loc[pod_df['WERKS'] == int(zplant), VehicleTypeCode].iloc[0])
        zshipcode = is_number(shipcode_df.loc[shipcode_df['ShipMode'] == VehicleTypeCode, 'ShipCode'].iloc[0])
        zpoo = is_number(poo_df.loc[poo_df['Factory'] == zfty, VehicleTypeCode].iloc[0])
        vendorCode = is_number(shipcode_df.loc[shipcode_df['ShipMode'] == VehicleTypeCode, 'LSPCode'].iloc[0])



    source =   {
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

    #print(f'source is {source["ScheduleLine"]}')
    source["shipment"]["deliveries"] = ttl_ship["deliveries"]
    source["shipment"]["goodsHolders"] = ttl_gh["goodsHolders"]
    st.write(source)
    jsource = json.dumps(source)
    #st.write(jsource)
    #r = requests.request("POST", base_url, headers=headers, data=jsource)
    #print(r.text)
    return None

###################create GUI##################

with st.form("Creat ASN"):
    st.write("Remember to download PO data into Testdata.xls")
    in_UUID = st.text_input("UUID")
    in_RID = st.text_input("Receipt ID")
    in_PO = st.text_input("CRPO number. Use , for more than one(no space)")
    in_TCCI = st.text_input("TCCI number")
    in_PDD = st.text_input("Planned Discharge Date like 2023-08-23")
    in_EDT = st.text_input("estimated Delivery Time stamp like 2023-07-31")
    in_FTY = st.selectbox("Factory", options , format_func=lambda x: dic[x])
    # Every form must have a submit button.
    submitted = st.form_submit_button("Create ASN from PO")
    if submitted:
        l_po = in_PO.split(",")
        st.session_state.UUID = in_UUID
        st.session_state.RID = in_RID
        st.session_state.TCCI = in_TCCI
        st.session_state.PDD = in_PDD
        st.session_state.EDT = str(in_EDT) + "T01:00:00Z"
        try:
            update_field(l_po, str(in_FTY))
        except Exception as E:
            st.write(E)
