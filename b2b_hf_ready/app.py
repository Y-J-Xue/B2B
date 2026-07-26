from __future__ import annotations

import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
from streamlit.components.v1 import html as components_html
from openpyxl import load_workbook

from converter import (
    ConversionConfig,
    InvoiceRow,
    build_workbook_from_data,
    extract_destination_country,
    load_source_data,
    save_workbook,
)


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "assets" / "template.xlsx"


def _cell_str(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def _cell_float(value):
    num = pd.to_numeric(value, errors="coerce")
    return None if pd.isna(num) else float(num)


def _field_label(label: str, required: bool = False) -> None:
    suffix = " <span style='color:red'>*</span>" if required else ""
    st.markdown(f"**{label}**{suffix}", unsafe_allow_html=True)


def _text_field(label: str, *, value: str = "", required: bool = False, **kwargs) -> str:
    _field_label(label, required=required)
    return st.text_input("", value=value, label_visibility="collapsed", **kwargs)


def _select_field(label: str, *, options, index: int = 0, required: bool = False, **kwargs):
    _field_label(label, required=required)
    return st.selectbox("", options, index=index, label_visibility="collapsed", **kwargs)


st.set_page_config(page_title="B2B下单模板生成器", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
[data-testid="stSidebar"] {display: none;}
</style>
""", unsafe_allow_html=True)

st.title("B2B 下单模板生成器")

components_html(
    """
    <script>
    (function() {
      function isEditableTarget(target) {
        if (!target) return false;
        const tag = (target.tagName || '').toLowerCase();
        return tag === 'input' || tag === 'textarea' || target.isContentEditable;
      }
      document.addEventListener('keydown', function(e) {
        const key = (e.key || '').toLowerCase();
        if ((e.ctrlKey || e.metaKey) && key === 'c' && !isEditableTarget(e.target)) {
          e.stopPropagation();
          e.preventDefault();
        }
      }, true);
      document.addEventListener('copy', function(e) {
        const active = document.activeElement;
        if (!isEditableTarget(active)) {
          e.stopPropagation();
        }
      }, true);
    })();
    </script>
    """,
    height=0,
)


if not TEMPLATE_PATH.exists():
    st.error(f"模板文件不存在：{TEMPLATE_PATH}")
    st.stop()

uploaded = st.file_uploader("上传原始 Excel（.xlsx）", type=["xlsx"])

input_path = None
source_ready = False
packing_rows_for_display = {}
invoice_rows_for_display = {}
summary_counts_for_display = {}
summary_reference_ids_for_display = {}


detected_country = ""

if uploaded is not None:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_in:
        tmp_in.write(uploaded.getvalue())
        input_path = Path(tmp_in.name)

    try:
        packing_rows_for_display, invoice_rows_for_display, summary_counts_for_display, summary_reference_ids_for_display = load_source_data(str(input_path))
        source_ready = True

        try:
            detected_country = extract_destination_country(str(input_path))
        except Exception:
            detected_country = ""
    except Exception as exc:
        st.error(f"无法读取原始文件：{exc}")
        st.stop()
    finally:
        try:
            if input_path is not None and input_path.exists():
                input_path.unlink(missing_ok=True)
        except Exception:
            pass


if uploaded is None:
    st.info("请先上传原始 Excel。")
    st.stop()


file_key = f"{uploaded.name}_{uploaded.size}"

st.divider()
st.subheader("模板待填项")

col1, col2 = st.columns(2)
with col1:
    product_code = _text_field(
        "产品代码",
        value="",
        required=True,
        placeholder="请输入产品代码",
        help="必填。",
        key=f"product_code_{file_key}",
    )
with col2:
    warehouse_code = _text_field(
        "仓库代码",
        value="",
        required=True,
        placeholder="例如 BNA6",
        help="必填。",
        key=f"warehouse_code_{file_key}",
    )

col3, col4 = st.columns(2)
with col3:
    destination_country = _text_field(
        "目的国家（二字码）",
        value=detected_country,
        required=True,
        placeholder="例如 US",
        help="必填。",
        key=f"destination_country_{file_key}",
    )
with col4:
    quantity_unit = _text_field(
        "数量单位",
        value="套",
        required=True,
        placeholder="例如 套",
        help="必填。",
        key=f"quantity_unit_{file_key}",
    )

col5, col6 = st.columns(2)
with col5:
    use_purpose = _text_field(
        "用途",
        value="服装",
        required=True,
        placeholder="默认 服装",
        help="必填。",
        key=f"use_purpose_{file_key}",
    )
with col6:
    currency = _text_field(
        "币种",
        value="USD",
        required=True,
        placeholder="默认 USD",
        help="必填。",
        key=f"currency_{file_key}",
    )

sales_link = _text_field(
    "销售链接",
    value="无",
    required=True,
    placeholder="默认 无",
    help="必填。",
    key=f"sales_link_{file_key}",
)

col7, col8 = st.columns(2)
with col7:
    default_magnet = _select_field(
        "是否带磁",
        options=["否", "是"],
        index=0,
        required=True,
        key=f"default_magnet_{file_key}",
    )
with col8:
    default_electric = _select_field(
        "是否带电",
        options=["否", "是"],
        index=0,
        required=True,
        key=f"default_electric_{file_key}",
    )

st.divider()
st.subheader("非必填项")

opt_rec1, opt_rec2, opt_rec3, opt_rec4 = st.columns(4)
with opt_rec1:
    recipient_name = _text_field(
        "收件人姓名",
        value="",
        required=False,
        placeholder="非必填",
        key=f"recipient_name_{file_key}",
    )
with opt_rec2:
    recipient_state = _text_field(
        "收件人省州",
        value="",
        required=False,
        placeholder="非必填",
        key=f"recipient_state_{file_key}",
    )
with opt_rec3:
    recipient_city = _text_field(
        "收件人城市",
        value="",
        required=False,
        placeholder="非必填",
        key=f"recipient_city_{file_key}",
    )
with opt_rec4:
    recipient_postal_code = _text_field(
        "收件人邮编",
        value="",
        required=False,
        placeholder="非必填",
        key=f"recipient_postal_code_{file_key}",
    )

opt0, opt00, opt000, opt001 = st.columns(4)
with opt0:
    recipient_address1 = _text_field(
        "收件人地址1",
        value="",
        required=False,
        placeholder="非必填",
        key=f"recipient_address1_{file_key}",
    )
with opt00:
    recipient_address2 = _text_field(
        "收件人地址2",
        value="",
        required=False,
        placeholder="非必填",
        key=f"recipient_address2_{file_key}",
    )
with opt000:
    recipient_phone = _text_field(
        "收件人电话",
        value="",
        required=False,
        placeholder="非必填",
        key=f"recipient_phone_{file_key}",
    )
with opt001:
    gross_weight = _text_field(
        "毛重KG",
        value="",
        required=False,
        placeholder="非必填",
        key=f"gross_weight_{file_key}",
    )

opt1, opt2, opt3 = st.columns(3)
with opt1:
    recipient_company = _text_field(
        "收件人公司",
        value="",
        required=False,
        placeholder="可选",
        key=f"recipient_company_{file_key}",
    )
with opt2:
    recipient_email = _text_field(
        "收件人邮箱",
        value="",
        required=False,
        placeholder="可选",
        key=f"recipient_email_{file_key}",
    )
with opt3:
    delivery_time = _text_field(
        "送达时段",
        value="",
        required=False,
        placeholder="可选",
        key=f"delivery_time_{file_key}",
    )

opt4, opt5, opt6 = st.columns(3)
with opt4:
    tax_id = _text_field(
        "税号/TAX ID",
        value="",
        required=False,
        placeholder="可选",
        key=f"tax_id_{file_key}",
    )
with opt5:
    importer_name = _text_field(
        "进口商名称",
        value="",
        required=False,
        placeholder="可选",
        key=f"importer_name_{file_key}",
    )
with opt6:
    bond_expiry = _text_field(
        "BOND有效期",
        value="",
        required=False,
        placeholder="可选",
        key=f"bond_expiry_{file_key}",
    )

opt7, opt8, opt9 = st.columns(3)
with opt7:
    eori = _text_field(
        "EORI",
        value="",
        required=False,
        placeholder="可选",
        key=f"eori_{file_key}",
    )
with opt8:
    importer_address = _text_field(
        "进口商地址",
        value="",
        required=False,
        placeholder="可选",
        key=f"importer_address_{file_key}",
    )
with opt9:
    report_method = _text_field(
        "报关方式",
        value="",
        required=False,
        placeholder="可选",
        key=f"report_method_{file_key}",
    )

opt10, opt11, opt12 = st.columns(3)
with opt10:
    insurance_service = _text_field(
        "保价服务",
        value="BJFDR",
        required=False,
        placeholder="默认 BJFDR",
        key=f"insurance_service_{file_key}",
    )
with opt11:
    signature_service = _text_field(
        "签名服务",
        value="",
        required=False,
        placeholder="可选",
        key=f"signature_service_{file_key}",
    )
with opt12:
    additional_service = _text_field(
        "附加服务",
        value="",
        required=False,
        placeholder="可选",
        key=f"additional_service_{file_key}",
    )

opt13, opt14, opt15 = st.columns(3)
with opt13:
    appointment_link = _text_field(
        "预约链接",
        value="",
        required=False,
        placeholder="可选",
        key=f"appointment_link_{file_key}",
    )
with opt14:
    appointment_code = _text_field(
        "预约码",
        value="",
        required=False,
        placeholder="可选",
        key=f"appointment_code_{file_key}",
    )
with opt15:
    remarks = _text_field(
        "备注",
        value="",
        required=False,
        placeholder="可选",
        key=f"remarks_{file_key}",
    )

opt16, opt17, opt18 = st.columns(3)
with opt16:
    sender_name = _text_field(
        "发件人姓名",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_name_{file_key}",
    )
with opt17:
    sender_company = _text_field(
        "发件人公司",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_company_{file_key}",
    )
with opt18:
    sender_country = _text_field(
        "发件人国家",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_country_{file_key}",
    )

opt19, opt20, opt21 = st.columns(3)
with opt19:
    sender_state = _text_field(
        "发件人省州",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_state_{file_key}",
    )
with opt20:
    sender_city = _text_field(
        "发件人城市",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_city_{file_key}",
    )
with opt21:
    sender_address = _text_field(
        "发件人详细地址",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_address_{file_key}",
    )

opt22, opt23, opt24 = st.columns(3)
with opt22:
    sender_postal_code = _text_field(
        "发件人邮编",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_postal_code_{file_key}",
    )
with opt23:
    sender_phone = _text_field(
        "发件人电话",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_phone_{file_key}",
    )
with opt24:
    sender_email = _text_field(
        "发件人邮箱",
        value="",
        required=False,
        placeholder="可选",
        key=f"sender_email_{file_key}",
    )
if st.button("生成模板", type="primary"):
    product_code_value = product_code.strip()
    warehouse_code_value = warehouse_code.strip()
    destination_country_value = destination_country.strip().upper()
    recipient_name_value = recipient_name.strip()
    recipient_company_value = recipient_company.strip()
    recipient_address1_value = recipient_address1.strip()
    recipient_address2_value = recipient_address2.strip()
    recipient_state_value = recipient_state.strip()
    recipient_city_value = recipient_city.strip()
    recipient_postal_code_value = recipient_postal_code.strip()
    recipient_phone_value = recipient_phone.strip()
    recipient_email_value = recipient_email.strip()
    tax_id_value = tax_id.strip()
    importer_name_value = importer_name.strip()
    bond_expiry_value = bond_expiry.strip()
    eori_value = eori.strip()
    importer_address_value = importer_address.strip()
    report_method_value = report_method.strip()
    insurance_service_value = insurance_service.strip()
    signature_service_value = signature_service.strip()
    additional_service_value = additional_service.strip()
    appointment_link_value = appointment_link.strip()
    appointment_code_value = appointment_code.strip()
    sender_name_value = sender_name.strip()
    sender_company_value = sender_company.strip()
    sender_country_value = sender_country.strip()
    sender_state_value = sender_state.strip()
    sender_city_value = sender_city.strip()
    sender_address_value = sender_address.strip()
    sender_postal_code_value = sender_postal_code.strip()
    sender_phone_value = sender_phone.strip()
    sender_email_value = sender_email.strip()
    delivery_time_value = delivery_time.strip()
    gross_weight_value = gross_weight.strip()
    remarks_value = remarks.strip()
    quantity_unit_value = quantity_unit.strip()
    use_purpose_value = use_purpose.strip() or "服装"
    currency_value = currency.strip().upper() or "USD"
    sales_link_value = sales_link.strip()

    if not product_code_value:
        st.error("请填写产品代码。")
        st.stop()
    if not warehouse_code_value:
        st.error("请填写仓库代码。")
        st.stop()
    if not destination_country_value or len(destination_country_value) != 2 or not destination_country_value.isalpha():
        st.error("目的国家必须是两个英文字母，例如 US。")
        st.stop()
    if not quantity_unit_value:
        st.error("请填写数量单位。")
        st.stop()
    if not sales_link_value:
        st.error("请填写销售链接。")
        st.stop()

    # 直接使用原始 Shipment ID 作为客户单号；客户单号与 INVOICE 现已隐藏为自动填充逻辑。
    shipment_id_map = {}

    config = ConversionConfig(
        product_code=product_code_value,
        destination_country=destination_country_value,
        warehouse_code=warehouse_code_value,
        recipient_name=recipient_name_value,
        recipient_company=recipient_company_value,
        recipient_address1=recipient_address1_value,
        recipient_address2=recipient_address2_value,
        recipient_state=recipient_state_value,
        recipient_city=recipient_city_value,
        recipient_postal_code=recipient_postal_code_value,
        recipient_phone=recipient_phone_value,
        recipient_email=recipient_email_value,
        tax_id=tax_id_value,
        importer_name=importer_name_value,
        bond_expiry=bond_expiry_value,
        eori=eori_value,
        importer_address=importer_address_value,
        report_method=report_method_value,
        insurance_service=insurance_service_value,
        signature_service=signature_service_value,
        additional_service=additional_service_value,
        appointment_link=appointment_link_value,
        appointment_code=appointment_code_value,
        sender_name=sender_name_value,
        sender_company=sender_company_value,
        sender_country=sender_country_value,
        sender_state=sender_state_value,
        sender_city=sender_city_value,
        sender_address=sender_address_value,
        sender_postal_code=sender_postal_code_value,
        sender_phone=sender_phone_value,
        sender_email=sender_email_value,
        delivery_time=delivery_time_value,
        use_purpose=use_purpose_value,
        currency=currency_value,
        sales_link=sales_link_value,
        gross_weight=gross_weight_value,
        separate_customs="否",
        default_magnet=default_magnet,
        default_electric=default_electric,
        quantity_unit=quantity_unit_value,
        remarks=remarks_value,
        shipment_id_map=shipment_id_map,
    )

    with st.spinner("正在生成，请稍候..."):
        wb, packing_rows, out_rows, summary_counts = build_workbook_from_data(
            packing_rows_for_display,
            invoice_rows_for_display,
            summary_counts_for_display,
            summary_reference_ids_for_display,
            str(TEMPLATE_PATH),
            config,
        )
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp_out:
            output_path = Path(tmp_out.name)
        save_workbook(wb, str(output_path))
        output_bytes = output_path.read_bytes()
        try:
            output_path.unlink(missing_ok=True)
        except Exception:
            pass

    st.success(f"生成完成：{len(out_rows)} 行明细。")
    st.download_button(
        label="下载生成的模板文件",
        data=output_bytes,
        file_name="B2B下单模板_生成结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
