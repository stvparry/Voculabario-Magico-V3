import re,traceback
from pathlib import Path
import streamlit as st
from analyser import analyse,build_workbook,reading_profile,map_duolingo,map_qualification,LEVEL_DESCRIPTIONS
from document_reader import extract_files
st.set_page_config(page_title='Vocabulario Mágico V3',page_icon='📚',layout='centered')
st.markdown('''<style>.block-container{max-width:1000px;padding-top:1.1rem}.hero{padding:1.35rem 1.6rem;border-radius:0 0 14px 14px;background:linear-gradient(135deg,#17365D,#2F75B5);color:white;margin-bottom:1rem}.hero h1{margin:0}.stripe{height:10px;background:linear-gradient(to right,#AA151B 0 25%,#F1BF00 25% 75%,#AA151B 75% 100%)}div.stButton>button[kind="primary"],div.stDownloadButton>button{background:#F1BF00!important;color:#4A3200!important;font-weight:700!important}</style>''',unsafe_allow_html=True)
img=Path(__file__).parent/'assets'/'spanish_tiles.png'
if img.exists():st.image(str(img),use_container_width=True)
st.markdown('<div class="stripe"></div><div class="hero"><h1>📚 Vocabulario Mágico V3</h1><p>Upload documents, scans or phone photos and create a personalised Spanish learning pack.</p></div>',unsafe_allow_html=True)
mode=st.radio('Choose source',('Upload files or images','Paste text'),horizontal=True);meta={'filename':None,'type':'Pasted text','method':'Pasted text','ocr_confidence':None};text=''
if mode=='Upload files or images':
    ocr_mode=st.radio('OCR mode',('Auto','Force OCR','Off'),horizontal=True,help='Auto uses embedded text when available and OCRs image-only PDFs.')
    files=st.file_uploader('PDF, DOCX, TXT, JPG, PNG, WebP, HEIC or TIFF',type=['pdf','docx','txt','jpg','jpeg','png','webp','heic','heif','tif','tiff'],accept_multiple_files=True)
    if files:
        try:
            with st.spinner('Reading files and running OCR where needed...'):text,meta=extract_files(files,ocr_mode)
            st.success(f"Extracted {len(text.split()):,} words using {meta['method']}.")
            if meta.get('ocr_confidence') is not None:st.info(f"Approximate OCR confidence: **{meta['ocr_confidence']:.1f}%**. Review names, accents and punctuation.")
            with st.expander('Preview and edit extracted text'):text=st.text_area('Extracted text',text,height=380)
        except Exception as e:st.error(str(e))
else:text=st.text_area('Paste Spanish text',height=320)
title=st.text_input('Title or source (optional)')
st.subheader('Learner level');system=st.radio('Level system',('Duolingo Score','CEFR band','UK qualification'),horizontal=True)
if system=='Duolingo Score':score=st.slider('Duolingo Score',0,160,105);cefr=map_duolingo(score);label=f'Duolingo Score {score} → {cefr}'
elif system=='CEFR band':cefr=st.select_slider('CEFR',['A1','A2','B1','B2','C1','C2'],value='B1');label=f'CEFR {cefr} ({LEVEL_DESCRIPTIONS[cefr]})'
else:q=st.selectbox('Qualification',('GCSE Foundation / early GCSE','GCSE Higher / strong GCSE','AS Level','A Level'));cefr=map_qualification(q);label=f'{q} → indicative {cefr}'
count=st.slider('Focus words',10,60,30,5);valid=len(re.findall(r'[A-Za-zÁÉÍÓÚÜÑáéíóúüñ]+',text))>=30
if text and not valid:st.warning('At least about 30 recognised words are needed.')
if st.button('Create personalised workbook',type='primary',disabled=not valid,use_container_width=True):
    try:
        result=analyse(text,cefr,label,count);profile=reading_profile(text,cefr);book=build_workbook(result,title,text,meta,profile)
        st.success(f"Ready: {len(result.words)} focus words; estimated {profile['estimated_cefr']} ({profile['challenge']}).")
        st.download_button('📥 Download workbook',book,'vocabulario_magico_v3.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
    except Exception:st.error('Workbook generation failed.');st.code(traceback.format_exc())
st.caption('V3 • scanned-PDF and image OCR • no AI key required')
