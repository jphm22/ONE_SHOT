# -*- coding: utf-8 -*-
"""Auditoría de presentaciones ya generadas. No usa Dimensions ni modifica PPTX."""
import argparse,json,re
from pathlib import Path
from pptx import Presentation

def auditar(archivo):
    prs=Presentation(str(archivo))
    total_bases_cero=0;total_none=0;sin_graficos=[];con_un_grafico=[];slides=[]
    for num,slide in enumerate(prs.slides,1):
        charts=sum(bool(shape.has_chart) for shape in slide.shapes)
        tables=sum(bool(shape.has_table) for shape in slide.shapes)
        textos='\n'.join(sh.text for sh in slide.shapes if sh.has_text_frame)
        bases_cero=len(re.findall(r'Base válida \(N\):\s*0\b',textos))
        none=sum('None' in sh.text for sh in slide.shapes if sh.has_text_frame and 'Valores absolutos' in sh.text)
        total_bases_cero+=bases_cero;total_none+=none
        if charts+tables==0:sin_graficos.append(num)
        elif charts+tables==1:con_un_grafico.append(num)
        slides.append({'slide':num,'graficos':charts,'tablas':tables,'paneles_N_cero':bases_cero,'paneles_None':none})
    return {'pptx':str(archivo),'diapositivas':len(prs.slides),'sin_graficos':sin_graficos,
            'con_un_grafico':con_un_grafico,'paneles_N_cero':total_bases_cero,
            'paneles_None':total_none,'diapositivas_detalle':slides}

def main():
    parser=argparse.ArgumentParser(description='Diagnóstico independiente del PPTX')
    parser.add_argument('--pptx',type=Path,required=True)
    parser.add_argument('--salida',type=Path)
    a=parser.parse_args()
    r=auditar(a.pptx)
    destino=a.salida or a.pptx.with_suffix('.auditoria.json')
    destino.write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
    print('Slides:',r['diapositivas'],'Sin gráficos:',r['sin_graficos'],
          'N=0:',r['paneles_N_cero'],'Conteos None:',r['paneles_None'])
    print('Auditoría:',destino)
if __name__=='__main__':main()
