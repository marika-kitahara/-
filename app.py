import io
import pandas as pd
import streamlit as st

st.set_page_config(page_title='アフィリエイト月次集計', layout='wide')
st.title('アフィリエイト月次集計')
st.caption('複数CSVを合算し、アフィリエイトID × 成果発生月で集計します。')
files = st.file_uploader('CSVファイルをアップロード（複数選択可）', type=['csv'], accept_multiple_files=True)
REQUIRED = ['成果承認フラグ','アフィリエイトID','パートナーサイト名','成果発生日時','グロス']

def read_csv(f):
    raw=f.getvalue()
    for enc in ('cp932','utf-8-sig','utf-8'):
        try:
            df=pd.read_csv(io.BytesIO(raw), encoding=enc, dtype={'アフィリエイトID':'string'}, low_memory=False)
            return df
        except UnicodeDecodeError:
            continue
    raise ValueError('対応する文字コードで読み込めませんでした')

if files:
    dfs=[]; errors=[]
    for f in files:
        try:
            d=read_csv(f)
            missing=[c for c in REQUIRED if c not in d.columns]
            if missing: raise ValueError('必須列がありません: '+', '.join(missing))
            dfs.append(d[REQUIRED].copy())
        except Exception as e: errors.append(f'{f.name}: {e}')
    if errors:
        for e in errors: st.error(e)
        st.stop()
    df=pd.concat(dfs,ignore_index=True)
    df['成果発生日時']=pd.to_datetime(df['成果発生日時'],errors='coerce')
    df['グロス']=pd.to_numeric(df['グロス'].astype('string').str.replace(',','',regex=False),errors='coerce')
    bad_date=int(df['成果発生日時'].isna().sum())
    bad_cost=int(df['グロス'].isna().sum())
    if bad_date: st.warning(f'成果発生日時が無効な {bad_date:,} 行は集計から除外します。')
    if bad_cost: st.warning(f'グロスが数値として読めない {bad_cost:,} 行はグロス0として集計します。')
    df=df.dropna(subset=['成果発生日時']).copy()
    df['グロス']=df['グロス'].fillna(0)
    df['アフィリエイトID']=df['アフィリエイトID'].fillna('').astype(str).str.strip()
    df['パートナーサイト名']=df['パートナーサイト名'].fillna('').astype(str).str.strip()
    df['月']=df['成果発生日時'].dt.to_period('M')
    df['発行']=df['成果承認フラグ'].astype(str).str.strip().str.upper().eq('Y')
    df['発行グロス値']=df['グロス'].where(df['発行'],0)
    keys=['アフィリエイトID','パートナーサイト名','月']
    agg=df.groupby(keys,dropna=False).agg(申込件数=('発行','size'),発行件数=('発行','sum'),グロス=('グロス','sum'),発行グロス=('発行グロス値','sum')).reset_index()
    agg['申込CPA']=agg['グロス']/agg['申込件数'].replace(0,float('nan'))
    agg['発行CPA']=agg['発行グロス']/agg['発行件数'].replace(0,float('nan'))
    metrics=['申込件数','発行件数','グロス','発行グロス','申込CPA','発行CPA']
    months=sorted(agg['月'].unique())
    labels=[f'{m.year}/{m.month}' for m in months]
    tables={}
    for metric in metrics:
        pivot=agg.pivot_table(index=keys[:2],columns='月',values=metric,aggfunc='first',dropna=False)
        # dropna=False can generate nonexistent ID/site combinations; retain only observed pairs.
        pairs=agg[keys[:2]].drop_duplicates()
        pivot=pivot.reindex(pd.MultiIndex.from_frame(pairs)).reindex(columns=months)
        if metric not in ('申込CPA','発行CPA'):
            pivot=pivot.fillna(0)
        pivot.columns=labels
        tables[metric]=pivot.reset_index().sort_values(keys[:2],kind='stable').reset_index(drop=True)
    st.success(f'{len(files)} ファイル・{len(df):,} 行を集計しました。')
    selected=st.selectbox('表示する指標',metrics)
    st.dataframe(tables[selected],use_container_width=True,hide_index=True)
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine='openpyxl') as writer:
        for metric,table in tables.items():
            table.to_excel(writer,sheet_name=metric,index=False)
            ws=writer.sheets[metric]
            ws.freeze_panes='C2'
            ws.auto_filter.ref=ws.dimensions
            ws.column_dimensions['A'].width=22
            ws.column_dimensions['B'].width=36
            for col in ws.iter_cols(min_col=3):
                ws.column_dimensions[col[0].column_letter].width=16
                for cell in col[1:]:
                    cell.number_format='#,##0.00' if metric.endswith('CPA') else '#,##0.##'
    st.download_button('集計Excelをダウンロード',data=out.getvalue(),file_name='アフィリエイト月次集計.xlsx',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
else:
    st.info('CSVをアップロードすると集計結果が表示されます。')
