from pathlib import Path
import datetime as dt
import io
import json
import tempfile
import zipfile
import streamlit as st
from report_engine import INPUTS, month_add
from cloud_engine import refresh_report

st.set_page_config(page_title='CityKart | GM Visibility', page_icon='📊', layout='wide')
st.markdown("""<style>
.stApp{background:#F5F7FB}
.block-container{max-width:1440px;padding-top:1.6rem;padding-bottom:2rem}
h1,h2,h3,h4{color:#18324B;letter-spacing:-.025em}
h1{font-size:2.15rem!important;font-weight:750!important}
h3{font-size:1.15rem!important}
[data-testid="stSidebar"]{background:#FFF;border-right:1px solid #E3E9F0}
[data-testid="stSidebar"] .block-container{padding-top:1.5rem}
[data-testid="stFileUploader"]{border:1px dashed #CCD6E1;border-radius:10px;background:#FAFCFE}
[data-testid="stFileUploader"] section{padding:1rem}
[data-testid="stMetric"]{background:white;border:1px solid #E3E9F0;border-radius:12px;padding:14px 18px}
[data-testid="stMetricLabel"]{color:#65758A;font-size:.8rem}
[data-testid="stMetricValue"]{font-size:1.4rem;color:#18324B}
[data-baseweb="tab-list"]{gap:1rem;border-bottom:1px solid #DFE6EE}
[data-baseweb="tab"]{font-weight:600;padding-left:4px;padding-right:4px}
[data-testid="stVerticalBlockBorderWrapper"]>div{border-color:#E2E8F0!important;border-radius:12px}
[data-testid="stAppDeployButton"]{display:none}
.brand{font-size:11px;letter-spacing:2px;font-weight:750;color:#B55B18;margin-bottom:7px}
.subtitle{color:#68798D;font-size:15px;margin-bottom:24px;line-height:1.6}
.side-brand{font-size:20px;font-weight:800;color:#18324B;letter-spacing:-.5px}
.side-sub{font-size:11px;color:#75849A;letter-spacing:1.5px;margin-bottom:25px}
.note{padding:15px 18px;background:#EAF0F6;border-radius:10px;color:#486077;font-size:14px;line-height:1.6}
</style>""", unsafe_allow_html=True)
base = Path(__file__).parent
mime_xlsx = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
template_files = [
    ('GRC_DATA', base/'input_templates/GRC_DATA_Upload_Template.xlsx'),
    ('PPO_DATA', base/'input_templates/PPO_DATA_Upload_Template.xlsx'),
    ('GM_TNA', base/'input_templates/GM_TNA_Upload_Template.xlsx'),
    ('MNL_PPO', base/'input_templates/MNL_PPO_Upload_Template.xlsx'),
    ('Regular buy plan', base/'buy_plan_templates/Regular_Buy_Plan_Template_Oct-Dec_2026.xlsx'),
    ('Winter buy plan', base/'buy_plan_templates/Winter_Buy_Plan_Template_Oct-Dec_2026.xlsx'),
    ('Report metadata', base/'master_templates/Report_Metadata_Upload_Template.xlsx'),
]
template_lookup = dict(template_files)
def download_template(label, widget_key):
    path = template_lookup[label]
    if path.exists():
        st.download_button('Download template', path.read_bytes(), path.name, mime=mime_xlsx,
                           key=widget_key, use_container_width=True)

with st.sidebar:
    st.markdown('<div class="side-brand">CITYKART</div><div class="side-sub">GM PLANNING & OPERATIONS</div>', unsafe_allow_html=True)
    st.subheader('Report settings')
    month = st.date_input('Starting month', dt.date(2026,10,1),
                         help='Regular planning uses this month and the following two months.')
    month = month.replace(day=1)
    period = ' / '.join(month_add(month,i).strftime('%b %Y') for i in range(3))
    st.caption(period)
    yesterday = dt.datetime.now(dt.timezone(dt.timedelta(hours=5, minutes=30))).date()-dt.timedelta(days=1)
    asof = st.date_input('GRC as of date', yesterday,
                        help='Defaults to one day before today (India time). Receipts run from the starting month through this cutoff, limited to month end. The report filename uses this cutoff date.')
    tna = st.selectbox('TNA quantity', ['TNA_Q ', 'TNA_Q ®'],
                      help='Choose the quantity field populated in your TNA upload.')
    st.divider()
    sidebar_status = st.empty()
    st.divider()
    st.caption('Online edition · No Excel installation required')
    st.caption('Download XLSX with live report formulas and a value-pasted summary.')

st.markdown('<div class="brand">GENERAL MERCHANDISE</div>', unsafe_allow_html=True)
st.title('Visibility report workspace')
st.markdown('<div class="subtitle">Prepare your source data and buy plans, review upload readiness, then download the detailed report and management summary.</div>', unsafe_allow_html=True)
overview = st.empty()
source_tab, plan_tab, metadata_tab, generate_tab, templates_tab = st.tabs(
    ['Source data', 'Buy plans', 'Metadata', 'Generate & downloads', 'Templates & help'])

uploads = {}
with source_tab:
    st.subheader('Daily source data')
    st.caption('Four inputs are required. SCHEDULE DT is treated as Delivery Date in PPO_DATA.')
    mode = st.radio('Upload format', ['Four separate files', 'One workbook with all four sheets'], horizontal=True)
    if mode.startswith('One'):
        with st.container(border=True):
            st.markdown('#### Combined source workbook')
            st.caption('Include GRC_DATA, PPO_DATA, GM_TNA and MNL_PPO as named worksheets.')
            combined = st.file_uploader('Upload source workbook', type=['xlsx','xlsb','xls'], key='combined')
            if combined:
                uploads = {n:(combined,n) for n in INPUTS}
                st.caption('All four source sheets will be checked during generation.')
    else:
        descriptions = {
            'GRC_DATA':('Receipt data','Receipt quantities and Entry Date.'),
            'PPO_DATA':('Pending purchase orders','Pending quantities and SCHEDULE DT = Delivery Date.'),
            'GM_TNA':('Floor & in-transit stock','The TNA quantity selected in Report settings.'),
            'MNL_PPO':('Manual purchase orders','Manual order quantities; a header-only file is accepted.'),
        }
        columns = st.columns(2)
        for i,name in enumerate(INPUTS):
            with columns[i%2]:
                with st.container(border=True):
                    title,description = descriptions[name]
                    st.markdown('#### '+title)
                    st.caption(name+' · '+description)
                    f = st.file_uploader('Upload '+name, type=['xlsx','xlsb','xls'], key=name)
                    with st.expander('Worksheet name'):
                        sheet = st.text_input('Sheet to read', value=name, key=name+'_sheet',
                                              help='A single-sheet workbook is selected automatically.')
                    download_template(name, 'source_template_'+name)
                    if f: uploads[name] = (f,sheet)

plans = {}
with plan_tab:
    st.subheader('Latest approved buy plans')
    st.caption('Upload both Regular and Winter plans. Keep only categories included in your current plan.')
    for column,group in zip(st.columns(2),('REGULAR','WINTER')):
        with column:
            with st.container(border=True):
                st.markdown('#### '+group.title()+' buy plan')
                st.caption(period if group=='REGULAR' else month.strftime('%b %Y')+' + full Winter season')
                f = st.file_uploader('Upload '+group.title()+' buy plan', type=['xlsx','xlsb','xls'], key=group+'_BP')
                if f: plans[group] = f
                download_template(group.title()+' buy plan', 'plan_template_'+group)
                st.caption('Department and Article sheets in one workbook.')
    with st.expander('Plan fields and missing articles'):
        st.write('Regular: OCT_2026_BUY_PLAN, NOV_2026_BUY_PLAN and DEC_2026_BUY_PLAN. Winter: OCT_2026_BUY_PLAN and WINTER_BUY_PLAN.')
        st.write('Reports use articles present in your source data or buy plans. Input-only articles are added with plan 0. Old template-only articles are omitted. Fill numeric quantities for retained plan rows; header-only plan sheets are allowed.')
        if month!=dt.date(2026,10,1):
            st.info('Templates are for October–December 2026. Rename their monthly plan headers for the selected months and year.')

with metadata_tab:
    st.subheader('Merchants and category attributes')
    st.caption('Optional · Upload one workbook containing all three metadata sheets.')
    with st.container(border=True):
        metadata_file = st.file_uploader('Upload report metadata', type=['xlsx','xlsb','xls'], key='report_metadata')
        download_template('Report metadata', 'metadata_template')
        st.markdown("""| Metadata sheet | Matched by | Report fields |
|---|---|---|
| DIVISION_MERCHANT | Division | Merchant |
| ARTICLE_MASTER | Article within Division / Section / Department | Article Status; STH percentages in Regular only |
| DEPARTMENT_MASTER | Division / Section / Department | Department Attribute; STH percentages in Regular only; Sale Rank |""")
        st.caption('Sale Rank uses Department data in both report levels. Article reports have no Attribute column; Winter reports have no STH columns. Uploaded fields are pasted as values; unmatched metadata remains blank.')
    with st.expander('Advanced: updated reference workbook'):
        custom = st.file_uploader('Optional report template', type=['xlsx'], key='template',
                                  help='Keep existing report columns and rows starting at row 5. Used for newer formatting or category metadata.')
        st.caption('New articles are added from uploaded data automatically. An updated reference workbook is not required for that.')

ready = len(uploads)==4 and len(plans)==2 and asof>=month
with overview.container():
    for column,label,value,detail in zip(
            st.columns(4),
            ['Planning period','Source files','Buy plans','Metadata'],
            [' / '.join(month_add(month,i).strftime('%b').upper() for i in range(3)),
             str(len(uploads))+' / 4', str(len(plans))+' / 2', 'Uploaded' if metadata_file else 'Optional'],
            [str(month.year), 'Required', 'Required', 'Division · Department · Article']):
        with column:
            st.metric(label,value)
            st.caption(detail)
with sidebar_status.container():
    st.caption('UPLOAD READINESS')
    st.progress((len(uploads)+len(plans))/6)
    st.write('Ready to generate' if ready else str(len(uploads)+len(plans))+' of 6 required uploads')
    st.caption('GRC cutoff: '+asof.strftime('%d %b %Y'))

with generate_tab:
    st.subheader('Generate your reports')
    missing = [n for n in INPUTS if n not in uploads]+[g.title()+' buy plan' for g in ('REGULAR','WINTER') if g not in plans]
    if missing:
        st.info('Still needed: '+', '.join(missing)+'. Use Source data and Buy plans to complete the uploads.')
    elif asof<month:
        st.warning('GRC as-of date must be on or after the starting month.')
    else:
        st.success('Required uploads are ready. Dates, matching keys and report calculations will be checked online.')
    summary_enabled = st.checkbox('Include management summary', value=True,
                                  help='Buyer/division comparisons for Regular and Winter, plus Bottom 15 MCs.')
    st.caption('Period: '+period+' · GRC through '+asof.strftime('%d %b %Y'))
    if st.button('Generate visibility report', type='primary', disabled=not ready, use_container_width=True):
        st.session_state.pop('generated',None)
        try:
            with tempfile.TemporaryDirectory(prefix='gm_visibility_') as tmp:
                work=Path(tmp);paths={}
                for name,(uploaded,sheet) in uploads.items():
                    p=work/(name+Path(uploaded.name).suffix.lower());p.write_bytes(uploaded.getvalue());paths[name]=(p,sheet)
                template=base/'template/GM_Visibility_Template.xlsx'
                if custom:
                    template=work/'template.xlsx';template.write_bytes(custom.getvalue())
                name=f'GM_Visibility_Report_Till_{asof:%d-%b-%y}.xlsx'
                output=work/name
                with st.status('Generating and checking your reports…', expanded=True) as status:
                    plan_paths={}
                    for group,f in plans.items():
                        p=work/(group+'_BP'+Path(f.name).suffix.lower());p.write_bytes(f.getvalue());plan_paths[group]=p
                    metadata_path=None
                    if metadata_file:
                        metadata_path=work/('report_metadata'+Path(metadata_file.name).suffix.lower());metadata_path.write_bytes(metadata_file.getvalue())
                    summary_output=work/f'GM_Summary_{month:%b-%Y}_Till_{asof:%d-%b-%Y}.xlsx' if summary_enabled else None
                    result=refresh_report(template,paths,output,month,asof,tna,lambda msg:status.write(msg),
                                          buy_plans=plan_paths,metadata_path=metadata_path,summary_destination=summary_output)
                    status.update(label='Detailed report ready; summary needs attention' if result.get('summary',{}).get('error') else 'Reports generated and checked online',state='complete')
                summary_ok=summary_output is not None and summary_output.exists() and not result.get('summary',{}).get('error')
                st.session_state.generated={'name':name,'content':output.read_bytes(),'audit':result,
                                            'summary_name':summary_output.name if summary_ok else None,
                                            'summary_content':summary_output.read_bytes() if summary_ok else None}
        except Exception as e:
            st.error(str(e))
            st.caption('Check the named sheet or field, correct the input and generate again.')
    if 'generated' in st.session_state:
        g=st.session_state.generated;a=g['audit']
        st.divider()
        st.subheader('Last generated report')
        st.caption('Starting month: '+a['report_month']+' · GRC as of '+a['as_of'])
        report_card,summary_card=st.columns(2)
        with report_card:
            with st.container(border=True):
                st.markdown('#### Detailed visibility report')
                st.caption('Four Regular / Winter sheets · Imported values + live report calculations')
                st.download_button('Download Excel report',g['content'],g['name'],
                                   mime=mime_xlsx,
                                   type='primary',use_container_width=True)
        with summary_card:
            with st.container(border=True):
                st.markdown('#### Management summary')
                st.caption('Buyer & division views · Regular / Winter · Bottom 15 MCs')
                if g.get('summary_content'):
                    st.download_button('Download summary',g['summary_content'],g['summary_name'],
                                       mime=mime_xlsx,type='primary',use_container_width=True)
                elif a.get('summary',{}).get('error'):
                    st.warning('Summary unavailable. The detailed report can still be downloaded.')
                    st.caption(a['summary']['error'])
                else:st.caption('Select Include management summary and generate again to create this file.')
        warnings=a.get('warnings',[])
        if warnings:st.warning(str(len(warnings))+' data checks need review. Open the details below.')
        with st.expander('Data checks and import details'):
            st.write('Imported source rows')
            st.dataframe([{'Input':n,'Rows':v} for n,v in a['rows'].items()],hide_index=True,use_container_width=True)
            st.write('Output rows')
            st.dataframe([{'Report':n,'Rows':v} for n,v in a['output_rows'].items()],hide_index=True,use_container_width=True)
            for message in warnings:st.warning(message)
            if a['differences']:
                st.caption('Values changed from the reference workbook; calculations use your latest uploads.')
                st.dataframe(a['differences'],use_container_width=True,hide_index=True,height=300)
            st.download_button('Download validation notes',json.dumps(a,indent=2,default=str),
                               'GM_Refresh_Comparison.json',mime='application/json')
        st.caption('To inspect the downloaded workbook’s source data, right-click a sheet tab and choose Unhide.')

with templates_tab:
    st.subheader('Input template library')
    st.caption('Seven templates for your complete reporting workflow.')
    archive=io.BytesIO()
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as z:
        for label,path in template_files:
            if path.exists():z.writestr(path.name,path.read_bytes())
    st.download_button('Download all templates (ZIP)',archive.getvalue(),'GM_All_Input_Templates.zip',
                       mime='application/zip',key='all_input_templates',type='primary')
    columns=st.columns(2)
    for i,(label,path) in enumerate(template_files):
        with columns[i%2]:
            with st.container(border=True):
                st.markdown('**'+label+'**')
                download_template(label,'library_template_'+label)
    with st.expander('Calculation rules'):
        st.write('Output rows sort A–Z by Merchant, Division, Section, Department (Article sheets only), then OND Buy Plan (Regular) or Winter Buy Plan (Winter), largest to smallest within each group. October–December captions use OND.')
        st.write('SCHEDULE DT = Delivery Date. The first PPO month includes that month and all earlier pending deliveries. Later PPO months use their own date windows. A separate Delivery Date column is not needed.')
        st.write('GRC uses Entry Date from the starting month through the cutoff. TNA uses your selected quantity field. Article reports match Article within Division / Section / Department; Department reports match Department. Sale Rank always matches Department.')
        st.write('Source-only articles are included with plan 0. Uploaded source quantities, buy plans, OND Buy Plan, total system PPO and metadata are value-pasted. Fill rates, visibility, coverage and adjusted buy-plan calculations remain formulas. The download shows four output sheets with normalized inputs hidden. Each output sheet opens with one Merchant cell selected. Regenerate to refresh imported values.')
        st.write('Article reports exclude Attribute. Winter reports exclude both STH percentage fields. Summary percentages use capped quantity totals divided by the matching buy-plan totals. All three complete summary sheets are pasted as values with font size 8; percentages and quantities display without decimals. Bottom 15 selects the lowest first-month Visibility % among departments with Sale Rank 1–100 and a positive first-month plan; selection updates each time you generate.')
    with st.expander('Required source headers'):
        st.markdown("""| Input | Required columns |
|---|---|
| GRC_DATA | ENTRY DATE, DIVISION, SECTION, DEPARTMENT, ARTICLE NAME, GRC QTY |
| PPO_DATA | DIVISION, SECTION, DEPARTMENT, ARTCLE NAME, PENDING_Q, SCHEDULE DT (or DELIVERY DATE) |
| GM_TNA | DIVISION, SECTION, DEPARTMENT, ARTICLE_NAME, selected TNA_Q |
| MNL_PPO | DIVISION, SECTION, DEPARTMENT, ARTICLE_NAME, TTL ORDER QTY |""")
        st.caption('Column order may vary. Keep category labels consistent; enter numeric quantities and Excel dates.')
st.caption('CityKart · General Merchandise Planning')
