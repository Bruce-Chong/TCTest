
import streamlit as st
import pandas as pd
import os
from pathlib import Path
import openpyxl
from datetime import datetime
import requests
import json
import xml.etree.ElementTree as ET

st.set_page_config(page_title='Create TCCI')
st.title('Create TCCI')
st.subheader('Choose an action')


##################retrieve info from excel ***************************

# Path Setting:
try:
    current_path = Path(__file__).parent.absolute()  # Get the file path for this py file
except:
    current_path = (Path.cwd())

# import the inv_sheet from the excel file
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
client_secret = "ahfkdDlcyaQDFByMFp1EkYj8PEGDq8yDD-7D8kPc15CM_AzOESkBeb5e94Sn3T9S"
grant_type = "client_credentials"
tokendata = {
    "grant_type": grant_type,
    "client_id": client_id,
    "client_secret": client_secret
}

#get token
nike_auth_url = "https://nike-qa.oktapreview.com/oauth2/ausa0mcornpZLi0C40h7/v1/token"
auth_response = requests.post(nike_auth_url, data=tokendata)
token = json.loads(auth_response.text)['access_token']
#print (token)
contentype = "'Content-Type': 'application/json'"
headers = { 'Authorization' : 'Bearer ' +token , 'Content-type': contentype}

######Set today's date###############
today_dt = datetime.now()
today_dt = today_dt.strftime("%Y-%m-%d")


def is_number(s):
  if(s is None):
    return None
  try:
    float(s)
    return str(int(s))
  except ValueError:
    return str(s)

def set_linesource(zmatnr, zpo, zitm, zfci,zuom, zsubline):
    line_source = """                           <LineItem Key=\"RID-{0}_{1}\">
                        \r\n                                        <Attribute AttributeTypeCd=\"CC\">OR</Attribute>
                        \r\n                                        <Attribute AttributeTypeCd=\"PL\">00100</Attribute>
                        \r\n                                        <Attribute AttributeTypeCd=\"BP\">{3}</Attribute>
                        \r\n                                        <Attribute AttributeTypeCd=\"SW\">01000</Attribute>
                        \r\n                                        <Reference SourceRefTypeCd=\"128\" RefTypeCd=\"BAF\">{0}</Reference>
                        \r\n                                        <Reference SourceRefTypeCd=\"128\" RefTypeCd=\"4H\">{3}</Reference>
                        \r\n                                        <Reference SourceRefTypeCd=\"128\" RefTypeCd=\"4B\">JO</Reference>
                        \r\n                                        <Date TimeZone=\"UTC\" DateTypeCd=\"003\">{4}</Date>
                        \r\n                                        <Measure Qualifier=\"SQ\" SourceQualifier=\"738\" SourceUOMCd=\"355\" UOMCd=\"PR\">{1}</Measure>
                        \r\n                                        <Measure Qualifier=\"QUR\" SourceQualifier=\"738\" SourceUOMCd=\"355\" UOMCd=\"CTN\">1</Measure>
                        \r\n                                        <Measure Qualifier=\"N\" SourceQualifier=\"738\" SourceUOMCd=\"355\" UOMCd=\"KG\">0</Measure>
                        \r\n                                        <Measure Qualifier=\"VOL\" SourceQualifier=\"738\" SourceUOMCd=\"355\" UOMCd=\"CR\">0</Measure>
                        \r\n                                        <Measure Qualifier=\"OR\" SourceUOMCd=\"355\" UOMCd=\"{5}\"></Measure>
                        \r\n{6}
                        \r\n                                    </LineItem>""".format(zpo, zitm, zmatnr,zfci, today_dt, zuom, zsubline)

    return line_source
def post_api(po, ctrl_no, uuid, fci, zfty):
    # url and data definition
    base_url = "https://api-inboundlogistics-test.nike.com/OriginReceipt/v1"

    MASTER_ITM = ""
    ttl_line = ""
    ttl_subline = ""
    MATNR = ""
    UOM = ""

    for row in ekpo_sheet.iter_rows():
        EBELN = row[0].value
        EBELP = is_number(row[1].value)
        TXZ01 = is_number(row[6].value)
        TXZ01.replace(" ", "")
        size = TXZ01.split(",", 1)
        if str(EBELN) == po:
            if EBELP[-2:] == "00":
                # only do this if there is more than one master item
                if MASTER_ITM != "":
                    ttl_line = ttl_line + set_linesource(MATNR, po, MASTER_ITM, fci, UOM, ttl_subline)
                    ttl_subline = ""
                MASTER_ITM = EBELP.zfill(5)
                MATNR = is_number(row[7].value)
                UOM = is_number(row[18].value)

            else:
                QTY = is_number(row[17].value)
                subline_source = """\r\n                                <Subline Key=\"RID-{0}_{1}{2}\">
                        \r\n                                            <Attribute AttributeTypeCd=\"IZ\">{2}</Attribute>
                        \r\n                                            <Measure Qualifier=\"OR\" SourceUOMCd=\"355\" UOMCd=\"{3}\">{4}</Measure>
                        \r\n                                        </Subline>""".format(str(po), MASTER_ITM, size[1].replace(" ", ""), UOM, QTY )
                if ttl_subline == "":
                    ttl_subline = subline_source
                else:
                    ttl_subline = ttl_subline + subline_source
    if ttl_line != "":
        ttl_line = ttl_line + set_linesource(MATNR, po, MASTER_ITM, fci, UOM, ttl_subline)
    else:
        ttl_line = set_linesource(MATNR, po, MASTER_ITM, fci, UOM, ttl_subline)

    VehicleTypeCode = ekpo_df.loc[ekpo_df['EBELN'] == int(po), 'EVERS'].iloc[0]
    zplant = is_number(ekpo_df.loc[ekpo_df['EBELN'] == int(po), 'WERKS'].iloc[0])
    zpoo = is_number(poo_df.loc[poo_df['Factory'] == zfty, VehicleTypeCode].iloc[0])
    vendorCode = is_number(shipcode_df.loc[shipcode_df['ShipMode'] == VehicleTypeCode, 'LSPCode'].iloc[0])

    source = """<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>        
            \r\n<SOAP-ENV:Envelope xmlns:SOAP-ENV=\"http://schemas.xmlsoap.org/soap/envelope/\" SOAP-ENV:encodingStyle=\"http://www.w3.org/2001/12/soap-encoding\">
            \r\n    <SOAP-ENV:Body>
            \r\n        <XMLBundle GeneratedBy=\"NKE_OUTBOUND_OR\" TrackID=\"95585806\">
            \r\n            <XMLTransmission CtrlNumber=\"{0}\" Receiver=\"NKE\" Sender=\"NIKETRADET\" SourceOwner=\"FHRCLNT300\" Timestamp=\"20230213 234334\">
            \r\n                <XMLGroup CtrlNumber=\"904336\" GroupType=\"BP\" IncludedMessages=\"1\">
            \r\n                    <XMLTransaction CtrlNumber=\"RID-{1}-{2}\" TransactionType=\"BPM-861\">
            \r\n                        <BpMessage MessageType=\"861\" PurposeCd=\"00\">
            \r\n                            <Mode>{3}</Mode>
            \r\n                            <Reference SourceRefTypeCd=\"128\" RefTypeCd=\"06\">FHRCLNT300</Reference>
            \r\n                            <Date TimeZone=\"UTC+7\" DateTypeCd=\"922\">{4} 1211</Date>
            \r\n                            <Location LocTypeCd=\"RL\">
            \r\n                                <LocationID Qualifier=\"UN\">{5}</LocationID>
            \r\n                            </Location>
            \r\n                            <TradePartner RoleCd=\"FW\">
            \r\n                                <TradePartnerID Qualifier=\"93\">{6}</TradePartnerID>
            \r\n                            </TradePartner>
            \r\n                            <TradePartner RoleCd=\"16\">
            \r\n                                <TradePartnerID Qualifier=\"93\">{7}</TradePartnerID>
            \r\n                            </TradePartner>
            \r\n                            <Document Key=\"CONF\" DocType=\"CONF\">
            \r\n                                <DocumentID>RID-{1}</DocumentID>
            \r\n                                <Reference SourceRefTypeCd=\"128\" RefTypeCd=\"IO\">F</Reference>
            \r\n                                <Reference SourceRefTypeCd=\"128\" RefTypeCd=\"LP\">CC</Reference>
            \r\n                                <Order Key=\"{1}\" OrderType=\"PO\">
            \r\n                                    <OrderID>{1}</OrderID>
            \r\n{8}
            \r\n                                </Order>
            \r\n                            </Document>
            \r\n                        </BpMessage>
            \r\n                    </XMLTransaction>
            \r\n                </XMLGroup>
            \r\n            </XMLTransmission>
            \r\n        </XMLBundle>
            \r\n    </SOAP-ENV:Body>
            \r\n</SOAP-ENV:Envelope>""".format(uuid, po, ctrl_no, VehicleTypeCode, today_dt, zpoo, vendorCode, zplant, ttl_line)

    st.write(source)
    #r = requests.request("POST", base_url, headers=headers, data=jsource)
    #print(r.text)
    return None

###################create GUI##################

with st.form("Create OR"):
    st.write("Remember to download PO data into Testdata.xls")
    in_UUID = st.text_input("UUID")
    in_ctrl = st.text_input("Transmission control number", value="0001")
    in_PO = st.text_input("PO number")
    in_FCI = st.text_input("FCI number")
    in_FTY = st.selectbox("Factory", options, format_func=lambda x: dic[x])

    # Every form must have a submit button.
    submitted = st.form_submit_button("Create OR for PO")
    if submitted:
        try:
            post_api(str(in_PO), str(in_ctrl), str(in_UUID), str(in_FCI), in_FTY)
        except Exception as E:
            st.write(E)
        #st.write("TCCI posted for PO number " + str(in_PO))