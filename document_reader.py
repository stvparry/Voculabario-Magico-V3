from io import BytesIO
from pathlib import Path
import fitz,pytesseract
from docx import Document
from PIL import Image,ImageEnhance,ImageFilter,ImageOps
from pillow_heif import register_heif_opener
from pypdf import PdfReader
from pytesseract import Output
register_heif_opener()
SUPPORTED={'.pdf','.docx','.txt','.jpg','.jpeg','.png','.webp','.heic','.heif','.tif','.tiff'}
def prep(im):
    im=ImageOps.exif_transpose(im).convert('RGB')
    if max(im.size)>2800:im.thumbnail((2800,2800),Image.Resampling.LANCZOS)
    return ImageEnhance.Contrast(ImageOps.grayscale(im)).enhance(1.35).filter(ImageFilter.SHARPEN)
def ocr(im):
    im=prep(im)
    try:data=pytesseract.image_to_data(im,lang='spa+eng',config='--oem 1 --psm 6',output_type=Output.DICT)
    except pytesseract.TesseractError:data=pytesseract.image_to_data(im,lang='eng',config='--oem 1 --psm 6',output_type=Output.DICT)
    lines=[];current=[];last=None;scores=[]
    for i,t in enumerate(data.get('text',[])):
        t=(t or '').strip()
        if not t:continue
        key=(data['block_num'][i],data['par_num'][i],data['line_num'][i])
        if last is not None and key!=last and current:lines.append(' '.join(current));current=[]
        current.append(t);last=key
        try:
            n=float(data['conf'][i]);scores.append(n) if n>=0 else None
        except:pass
    if current:lines.append(' '.join(current))
    return '\n'.join(lines),round(sum(scores)/len(scores),1) if scores else 0
def one(f,mode):
    name=f.name;ext=Path(name).suffix.lower();raw=f.getvalue()
    if ext not in SUPPORTED:raise ValueError(f'Unsupported type {ext}')
    if ext=='.txt':
        for enc in ('utf-8-sig','utf-8','latin-1'):
            try:return raw.decode(enc),{'filename':name,'type':'TXT','method':'Text decode','ocr_confidence':None,'pages':0}
            except UnicodeDecodeError:pass
    if ext=='.docx':
        doc=Document(BytesIO(raw));text='\n'.join(p.text for p in doc.paragraphs if p.text.strip());return text,{'filename':name,'type':'DOCX','method':'DOCX extraction','ocr_confidence':None,'pages':0}
    if ext=='.pdf':
        reader=PdfReader(BytesIO(raw));digital='\n\n'.join((p.extract_text() or '').strip() for p in reader.pages);need=mode=='Force OCR' or (mode=='Auto' and len(digital)<max(120,35*len(reader.pages)))
        if not need or mode=='Off':return digital,{'filename':name,'type':'PDF','method':'Embedded PDF text','ocr_confidence':None,'pages':len(reader.pages)}
        doc=fitz.open(stream=raw,filetype='pdf');parts=[];scores=[]
        for page in doc:
            pix=page.get_pixmap(dpi=220,alpha=False);text,score=ocr(Image.frombytes('RGB',(pix.width,pix.height),pix.samples));parts.append(text);scores.append(score)
        return '\n\n'.join(parts),{'filename':name,'type':'Scanned PDF','method':'Tesseract OCR','ocr_confidence':round(sum(scores)/len(scores),1) if scores else 0,'pages':len(doc)}
    text,score=ocr(Image.open(BytesIO(raw)));return text,{'filename':name,'type':'Image','method':'Tesseract OCR','ocr_confidence':score,'pages':1}
def extract_files(files,mode='Auto'):
    texts=[];metas=[]
    for f in files:
        text,meta=one(f,mode);texts.append(f'--- {f.name} ---\n{text}');metas.append(meta)
    scores=[m['ocr_confidence'] for m in metas if m.get('ocr_confidence') is not None]
    return '\n\n'.join(texts),{'filename':', '.join(m['filename'] for m in metas),'type':'Multiple files' if len(metas)>1 else metas[0]['type'],'method':', '.join(sorted(set(m['method'] for m in metas))),'ocr_confidence':round(sum(scores)/len(scores),1) if scores else None,'pages':sum(m['pages'] for m in metas)}
