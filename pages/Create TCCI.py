
import streamlit as st
import pandas as pd
import os
from pathlib import Path
import openpyxl
from datetime import datetime
import requests
import json

st.set_page_config(page_title='Create TCCI')
st.title('Create TCCI')
st.subheader('Choose an action')


##################retrieve info from excel ***************************
# import the inv_sheet from the excel file
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
# put worksheet into dataframe
shipcode_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='ship_code', keep_default_na=False)
poo_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='POO', keep_default_na=False)
pod_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='POD', keep_default_na=False)
plant_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='plant_add', keep_default_na=False)
fac_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='factory', keep_default_na=False)
ekpo_df = pd.read_excel(filepath, engine='openpyxl', sheet_name='EKPO', keep_default_na=False)

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
##auth_response = requests.post(nike_auth_url, data=tokendata)
##token = json.loads(auth_response.text)['access_token']
#print (token)
contentype = "'Content-Type': 'application/json'"
##headers = { 'Authorization' : 'Bearer ' +token , 'Content-type': contentype}

######Set today's date###############
today_dt = datetime.now()
today_dt = today_dt.strftime("%Y-%m-%d")

###############set all session state variables###############
st.session_state.today_dt = str(today_dt)
st.session_state.crpo = ""
st.session_state.fci = ""
st.session_state.schedule_line = ""
st.session_state.CRPOLineItemNbr = ""
st.session_state.POLineItemNbr = ""
st.session_state.SizeDesc = ""
st.session_state.SchedLineReqCat = ""
st.session_state.OrderQty = ""
st.session_state.CancelDt = ""
st.session_state.DestinationPlant  = ""
st.session_state.CustomerInvoiceCreateDt = today_dt
st.session_state.ProductCd = ""
st.session_state.RequirementCategoryCd = ""
st.session_state.SalesUOMCd = ""
st.session_state.SchedLineNetNetWeight = ""
st.session_state.SchedLineNetNetSumWeight = ""
st.session_state.jschedule_line = ""
st.session_state.mode_of_transport = ""
st.session_state.vessel_airline = ""
st.session_state.shipped_from = ""
st.session_state.shipped_to = ""
st.session_state.letter_of_credit = ""
st.session_state.footer_text = ""
st.session_state.total_netweight = ""
st.session_state.total_gross_weight = ""
st.session_state.sys = ""
st.session_state.dmr = ""
st.session_state.HeaderNotationsName = ""
st.session_state.HeaderNotationsAddr1 = ""
st.session_state.HeaderNotationsAddr2 = ""
st.session_state.HeaderNotationsAddr3 = ""
st.session_state.HeaderNotationsAddr4 = ""
st.session_state.HeaderNotationsAddr5 = ""
st.session_state.HeaderNotationsAddr6 = ""
st.session_state.HeaderNotationsAddr7 = ""
st.session_state.AddressNm = ""
st.session_state.AddressTxt = ""
st.session_state.RegionCd = ""
st.session_state.RegionTxt = ""
st.session_state.CountryCd = ""
st.session_state.CountryTxt = ""
st.session_state.CityNm = ""
st.session_state.PostalCd = ""

def is_number(s):
  if(s is None):
    return None
  try:
    float(s)
    return str(int(s))
  except ValueError:
    return str(s)


def set_session(ypo, yfty):
    st.session_state.DestinationPlant = is_number(ekpo_df.loc[ekpo_df['EBELN'] == int(ypo), 'WERKS'].iloc[0])
    st.session_state.mode_of_transport = ekpo_df.loc[ekpo_df['EBELN'] == int(ypo), 'EVERS'].iloc[0]
    st.session_state.total_netweight = is_number(ekpo_df.loc[ekpo_df['EBELN'] == int(ypo), 'NETWR'].iloc[0])
    st.session_state.total_gross_weight = is_number(ekpo_df.loc[ekpo_df['EBELN'] == int(ypo), 'BRTWR'].iloc[0])
    try:
        st.session_state.AddressNm = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'AddressNm'].iloc[0])
        st.session_state.AddressTxt = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'AddressNm'].iloc[0])
        st.session_state.RegionCd = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'RegionCd'].iloc[0])
        st.session_state.RegionTxt = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'RegionTxt'].iloc[0])
        st.session_state.CountryCd = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'CountryCd'].iloc[0])
        st.session_state.CountryTxt = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'CountryTxt'].iloc[0])
        st.session_state.CityNm = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'CityNm'].iloc[0])
        st.session_state.PostalCd = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'PostalCd'].iloc[0])
        st.session_state.shipped_to = is_number(
            plant_df.loc[plant_df['WERKS'] == int(st.session_state.DestinationPlant), 'CountryCd'].iloc[0])
    except:
        st.write("Error with plant data for PO, plant is " + str(st.session_state.DestinationPlant))

    try:
        st.session_state.HeaderNotationsName = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsName'].iloc[0])
        st.session_state.HeaderNotationsAddr1 = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsAddr1'].iloc[0])
        st.session_state.HeaderNotationsAddr2 = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsAddr2'].iloc[0])
        st.session_state.HeaderNotationsAddr3 = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsAddr3'].iloc[0])
        st.session_state.HeaderNotationsAddr4 = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsAddr4'].iloc[0])
        st.session_state.HeaderNotationsAddr5 = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsAddr5'].iloc[0])
        st.session_state.HeaderNotationsAddr6 = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsAddr6'].iloc[0])
        st.session_state.HeaderNotationsAddr7 = is_number(
            fac_df.loc[fac_df['vendor'] == yfty, 'HeaderNotationsAddr7'].iloc[0])
        st.session_state.shipped_from = is_number(fac_df.loc[fac_df['vendor'] == yfty, 'MCO'].iloc[0])
    except:
        st.write("Error with factory data for PO, factory is " + yfty)
def update_field(zpo, zfci, zfty, zplant):

    empty_po = {"Line": [{}]}
    empty_sl = {"ScheduleLine": [{}]}
    total_po = {"Line": [{}]}
    total_sl = {"ScheduleLine": [{}]}

    st.session_state.fci = is_number(zfci)
    #st.dataframe(ekpo_df)
    masteritm = "0"
    lineno = 0
    for row in ekpo_sheet.iter_rows():
        EBELN = row[0].value
        EBELP = is_number(row[1].value)
        TXZ01 = is_number(row[6].value)
        TXZ01.replace(" ", "")
        size = TXZ01.split(",", 1)
        if str(EBELN) in zpo:
            if EBELP[-2:] == "00":
                set_session(EBELN, zfty)
                jpo = {
                  "Line": [
                    {
                      "ProductCd": is_number(row[7].value),
                      "CancelDtSpecified": "false",
                      "CRDSpecified": "false",
                      "PlantCd": str(zplant),
                      "RequirementCategoryCd": "01000",
                      "SalesUOMCd": is_number(row[18].value),
                      "CountryRegionPO": str(EBELN),
                      "LineTxt": [
                        {
                          "languageCd": "EN",
                          "textId": 7,
                          "Value": "FIRST QUALITY"
                        },
                        {
                          "languageCd": "EN",
                          "textId": 1,
                          "Value": "material_content"
                        },
                        {
                          "languageCd": "EN",
                          "textId": 0,
                          "Value": TXZ01
                        },
                        {
                          "languageCd": "EN",
                          "textId": 4,
                          "Value": str(zplant)
                        },
                        {
                          "languageCd": "EN",
                          "textId": 6,
                          "Value": st.session_state.shipped_from
                        },
                        {
                          "languageCd": "EN",
                          "textId": 5,
                          "Value": st.session_state.shipped_to
                        },
                        {
                          "languageCd": "EN",
                          "textId": 8,
                          "Value": "5"
                        },
                        {
                          "languageCd": "EN",
                          "textId": 9,
                          "Value": is_number(row[17].value)
                        },
                        {
                          "languageCd": "EN",
                          "textId": 10,
                          "Value": is_number(row[27].value)
                        },
                        {
                          "languageCd": "EN",
                          "textId": 2,
                          "Value": "Standard Description"
                        },
                        {
                          "languageCd": "EN",
                          "textId": 11,
                          "Value": "textile_category"
                        },
                        {
                          "languageCd": "EN",
                          "textId": 12,
                          "Value": ""
                        },
                        {
                          "languageCd": "EN",
                          "textId": 13,
                          "Value": ""
                        },
                        {
                          "languageCd": "EN",
                          "textId": 14,
                          "Value": ""
                        }
                      ],
                      "ScheduleLine": [{}]
                    }
                  ]
                }
                # only do this if there is more than one master item
                if total_po == empty_po:
                    total_po["Line"] = jpo["Line"]
                else:
                    total_po["Line"][lineno]["ScheduleLine"] = total_sl["ScheduleLine"]
                    total_po["Line"] += jpo["Line"]
                    total_sl = {"ScheduleLine": [{}]}
                    lineno = lineno + 1

                masteritm = EBELP
            else:
                schedule_line = {
                  "ScheduleLine": [
                    {
                      "CRPOLineItemNbr": is_number(masteritm),
                      "POLineItemNbr": is_number(masteritm),
                      "SizeDesc": size[1].replace(" ", ""),
                      "SchedLineReqCat": "01000",
                      "OrderQty": is_number(row[17].value),
                      "SchedLineNetNetWeight": is_number(row[26].value),
                      "SchedLineNetNetSumWeight": is_number(row[27].value)
                    }
                  ]
                }
                if total_sl == empty_sl:
                    total_sl["ScheduleLine"] = schedule_line["ScheduleLine"]

                else:
                    total_sl["ScheduleLine"] += schedule_line["ScheduleLine"]


    total_po["Line"][lineno]["ScheduleLine"] = total_sl["ScheduleLine"]
    #st.write(total_po)

    post_api(total_po)
    return None

def post_api(ttl_po):
  # url and data definition
  base_url = "https://nikecfqaapiportal.prod.apimanagement.us20.hana.ondemand.com:443/delivery/v1/TCOutboundDelivery"

  source =   {
    "EventControl": {
        "messageActionCode": "01",
        "sourceSystemName": "Mercury",
        "recordCreatedOnDate": st.session_state.today_dt,
        "recordCreatedOnTime": "190945",
        "messageIdentifier": "19926345"
    },
    "InvoiceRequest": {
      "DMRNumber": "",
      "TypeCd": "ZLF2",
      "CustomerPOTypeCd": "FFS",
      "CancelDt": "2023-12-31T00:00:00",
      "CRD": st.session_state.today_dt,
      "AdditionalDt": "0000-00-00T00:00:00",
      "AdditionalDtSpecified": "true",
      "DestinationPlant": st.session_state.DestinationPlant,
      "CustomerPONbr": st.session_state.fci,
      "CustomerInvoiceCreateDt": st.session_state.today_dt,
      "ShipToAddress": {
        "AddressNm": [
                  {
                      "sequenceNbr": "1",
                      "Value": "Nike USA"
                  },
                  {
                      "sequenceNbr": "2",
                      "Value": st.session_state.AddressNm
                  }
              ],
              "AddressTxt": {
                  "sequenceNbr": "1",
                  "Value": st.session_state.AddressTxt
              },
              "RegionCd": st.session_state.RegionCd,
              "RegionTxt": st.session_state.RegionTxt,
              "CountryCd": st.session_state.CountryCd,
              "CountryTxt": st.session_state.CountryTxt,
              "CityNm": st.session_state.CityNm,
              "PostalCd": st.session_state.PostalCd
          },

      "HeaderNotationsName": st.session_state.HeaderNotationsName,
      "HeaderNotationsAddr1": st.session_state.HeaderNotationsAddr1,
      "HeaderNotationsAddr2": st.session_state.HeaderNotationsAddr2,
      "HeaderNotationsAddr3": st.session_state.HeaderNotationsAddr3,
      "HeaderNotationsAddr4": st.session_state.HeaderNotationsAddr4,
      "HeaderNotationsAddr5": st.session_state.HeaderNotationsAddr5,
      "HeaderNotationsAddr6": st.session_state.HeaderNotationsAddr6,
      "HeaderNotationsAddr7": st.session_state.HeaderNotationsAddr7,
      "HeaderTxt": [
        {
          "languageCd": "EN",
          "textId": 0,
          "Value": st.session_state.mode_of_transport
        },
        {
          "languageCd": "EN",
          "textId": 1,
          "Value": "Maersk"
        },
        {
          "languageCd": "EN",
          "textId": 3,
          "Value": st.session_state.shipped_from
        },
        {
          "languageCd": "EN",
          "textId": 4,
          "Value": st.session_state.shipped_to
        },
        {
          "languageCd": "EN",
          "textId": 5,
          "Value": "LETTER OF CREDIT"
        },
        {
          "languageCd": "EN",
          "textId": 2,
          "Value": "total carton for PO is 5"
        },
        {
          "languageCd": "EN",
          "textId": 7,
          "Value": "FOOTER TEXT"
        },
        {
          "languageCd": "EN",
          "textId": 9,
          "Value": st.session_state.total_netweight
        },
        {
          "languageCd": "EN",
          "textId": 10,
          "Value": st.session_state.total_gross_weight
        }
      ],
      "Line": [
        {}
      ]
    }
  }

  #print(f'source is {source["ScheduleLine"]}')
  source["InvoiceRequest"]["Line"] = ttl_po["Line"]
  st.write(source)
  jsource = json.dumps(source)
  #st.write(jsource)
  #r = requests.request("POST", base_url, headers=headers, data=jsource)
  #print(r.text)
  return None

###################create GUI##################

with st.form("Creat TCCI"):
    st.write("Remember to download PO data into Testdata.xls")
    in_PO = st.text_input("PO Data. Use , for more than one PO(no space)")
    in_FCI = st.text_input("FCI")
    in_Vendor = st.selectbox("Factory", options, format_func=lambda x: dic[x])
    in_plant = st.text_input("Plant code. 1074 or 1198")
    # Every form must have a submit button.
    submitted = st.form_submit_button("Create TCCI from PO")
    if submitted:
        l_po = in_PO.split(",")
        # Path Setting:
        try:
            current_path = Path(__file__).parent.absolute()  # Get the file path for this py file
        except:
            current_path = (Path.cwd())


        try:
            update_field(l_po, str(in_FCI), str(in_Vendor), in_plant)
        except Exception as E:
            st.write(E)
        #st.write("TCCI posted for PO number " + str(in_PO))

