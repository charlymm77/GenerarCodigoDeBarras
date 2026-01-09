import { Component, ElementRef, ViewChild } from '@angular/core';
import JsBarcode from 'jsbarcode';
import jsPDF from 'jspdf';
import * as XLSX from 'xlsx';
import QRCode from 'qrcode';
import JSZip from 'jszip';

type BarcodeType =
  | 'CODE128' | 'EAN13' | 'EAN8' | 'UPC' | 'ITF14' | 'MSI' | 'pharmacode' | 'codabar';

type LabelLayout = 'classic' | 'logoLeft' | 'codeTop';

type Align = 'left' | 'center' | 'right';

interface LabelConfig {
  value: string;
  type: BarcodeType;
  mode: 'barcode' | 'qr';
  showValue: boolean;
  description: string;
  price: string;
  lot: string;
  headerText: string;
  footerText: string;
  logoDataUrl?: string;
  labelWidthMm: number;
  labelHeightMm: number;
  marginMm: number;
  dpi: number;
  layout: LabelLayout;
  bodyAlign: Align;
  headerScale: number;
  bodyScale: number;
  footerScale: number;
  qty?: number;
}

@Component({
  selector: 'app-label-designer',
  standalone: false,
  templateUrl: './label-designer.html',
  styleUrl: './label-designer.css',
})
export class LabelDesigner {
  value = '123456789012';
  renderError: string | null = null;
  formError: string | null = null;
  type: BarcodeType = 'CODE128';
  mode: 'barcode' | 'qr' = 'barcode';
  showValue = true;

  description = '';
  price = '';
  lot = '';
  headerText = '';
  footerText = '';
  logoDataUrl?: string;

  // Presets de tamaño
  templates = [
    { name: '50 x 30 mm', w: 50, h: 30, m: 5 },
    { name: '70 x 25 mm', w: 70, h: 25, m: 4 },
    { name: '100 x 50 mm', w: 100, h: 50, m: 6 },
    { name: 'Custom', w: 0, h: 0, m: 0 },
  ];
  selectedTemplate = '50 x 30 mm';

  labelWidthMm = 50;
  labelHeightMm = 30;
  quantity = 1;
  pageFormat: 'a4' | 'letter' = 'a4';
  marginMm = 5;
  dpi = 300;

  // Diseño y tipografía
  layout: LabelLayout = 'classic';
  bodyAlign: Align = 'center';
  headerScale = 0.10; // proporción del alto de la etiqueta
  bodyScale = 0.08;
  footerScale = 0.07;

  // Carga por Excel
  batchItems: LabelConfig[] = [];
  batchErrors: string[] = [];
  excelRows: any[] = [];
  selectedBatchIndex: number | null = null;
  get batchHasErrors() { return this.batchErrors.some(e => !!e); }
  get batchErrorCount() { return this.batchErrors.filter(Boolean).length; }
  useBatch = false;

  // Mapeo de columnas personalizadas (vacío = usar nombres reconocidos)
  customMap: Record<string, string> = {
    Codigo: '', Tipo: '', Modo: '', Cantidad: '', Descripcion: '', Precio: '', Lote: '',
    Encabezado: '', Pie: '', Ancho_mm: '', Alto_mm: '', Margen_mm: '', DPI: '', Layout: '',
    Alineacion: '', EscalaEncabezado: '', EscalaCuerpo: '', EscalaPie: '', LogoUrl: '',
  };

  @ViewChild('previewCanvas', { static: true }) previewCanvas!: ElementRef<HTMLCanvasElement>;

  // Preferencias: confirmar antes de limpiar
  confirmOnResetForm = true;
  confirmOnResetMapping = true;

  ngOnInit() {
    try {
      const a = localStorage.getItem('confirmOnResetForm');
      const b = localStorage.getItem('confirmOnResetMapping');
      if (a !== null) this.confirmOnResetForm = a !== 'false';
      if (b !== null) this.confirmOnResetMapping = b !== 'false';
    } catch {}
  }

  ngAfterViewInit() {
    this.renderPreview();
  }

  onChange() {
    this.updateFormError();
    this.renderPreview();
  }

  onTemplateChange() {
    const t = this.templates.find(x => x.name === this.selectedTemplate);
    if (t && t.name !== 'Custom') {
      this.labelWidthMm = t.w;
      this.labelHeightMm = t.h;
      this.marginMm = t.m;
    }
    this.updateFormError();
    this.renderPreview();
  }

  async renderPreview() {
    const canvas = this.previewCanvas.nativeElement;
    const pxW = this.mmToPx(this.labelWidthMm, 96);
    const pxH = this.mmToPx(this.labelHeightMm, 96);
    canvas.width = Math.max(1, Math.floor(pxW));
    canvas.height = Math.max(1, Math.floor(pxH));

    if (this.useBatch && this.selectedBatchIndex != null && this.batchItems[this.selectedBatchIndex]) {
      await this.drawLabelToCanvasWith(this.batchItems[this.selectedBatchIndex], canvas, 96);
    } else {
      await this.drawLabelToCanvas(canvas, 96);
    }
  }

  async exportPdf() {
    const doc = new jsPDF({
      unit: 'mm',
      format: this.pageFormat,
      compress: true,
    });

    const page = this.getPageSizeMm(this.pageFormat);

    // Construir lista de etiquetas a imprimir (batch o la actual)
    let items: LabelConfig[] = [];
    if (this.useBatch && this.batchItems.length) {
      for (const it of this.batchItems) {
        const count = Math.max(1, Math.floor(it.qty || 1));
        for (let k = 0; k < count; k++) items.push(it);
      }
    } else {
      for (let i = 0; i < this.quantity; i++) items.push(this.getCurrentConfig());
    }

    const gap = this.marginMm;
    const labelW = this.labelWidthMm;
    const labelH = this.labelHeightMm;

    const cols = Math.max(1, Math.floor((page.w - gap) / (labelW + gap)));
    const rows = Math.max(1, Math.floor((page.h - gap) / (labelH + gap)));
    const perPage = cols * rows;

    // Filtrar inválidos
    const errors = this.useBatch ? this.batchErrors : [];

    for (let i = 0; i < items.length; i++) {
      if (errors[i] && this.useBatch) continue;
      const idx = i % perPage;
      if (i > 0 && idx === 0) {
        doc.addPage(this.pageFormat);
      }
      const col = idx % cols;
      const row = Math.floor(idx / cols);
      const x = gap + col * (labelW + gap);
      const y = gap + row * (labelH + gap);

      const imgData = await this.buildLabelDataUrlFrom(items[i], this.dpi);
      doc.addImage(imgData, 'PNG', x, y, labelW, labelH);
    }

    doc.save('etiquetas.pdf');
  }

  async printPage() {
    // Construye una página HTML en mm para imprimir directamente
    const page = this.getPageSizeMm(this.pageFormat);
    const gap = this.marginMm;
    const labelW = this.labelWidthMm;
    const labelH = this.labelHeightMm;
    const cols = Math.max(1, Math.floor((page.w - gap) / (labelW + gap)));
    const rows = Math.max(1, Math.floor((page.h - gap) / (labelH + gap)));
    const perPage = cols * rows;

    // Lista de etiquetas a imprimir (batch o actual)
    let items: LabelConfig[] = [];
    if (this.useBatch && this.batchItems.length) {
      for (const it of this.batchItems) {
        const count = Math.max(1, Math.floor(it.qty || 1));
        for (let k = 0; k < count; k++) items.push(it);
      }
    } else {
      for (let i = 0; i < this.quantity; i++) items.push(this.getCurrentConfig());
    }

    const htmlParts: string[] = [];
    htmlParts.push(`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Imprimir etiquetas</title>`);
    htmlParts.push(`<style>
      @page { size: ${this.pageFormat === 'a4' ? '210mm 297mm' : '216mm 279mm'}; margin: 0; }
      html, body { margin:0; padding:0; }
      .page { position: relative; width: ${page.w}mm; height: ${page.h}mm; }
      img { position: absolute; width: ${labelW}mm; height: ${labelH}mm; }
    </style></head><body>`);

    let printed = 0;
    const errors = this.useBatch ? this.batchErrors : [];
    while (printed < items.length) {
      htmlParts.push(`<div class="page">`);
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          if (printed >= items.length) break;
          if (errors[printed]) { printed++; c--; continue; }
          const x = gap + c * (labelW + gap);
          const y = gap + r * (labelH + gap);
          const imgData = await this.buildLabelDataUrlFrom(items[printed], this.dpi);
          htmlParts.push(`<img src="${imgData}" style="left:${x}mm; top:${y}mm;" />`);
          printed++;
        }
      }
      htmlParts.push(`</div>`);
    }

    htmlParts.push(`<script>window.onload = () => { window.print(); setTimeout(()=>window.close(), 200); };</script></body></html>`);
    const w = window.open('', '_blank');
    if (w) {
      w.document.open();
      w.document.write(htmlParts.join(''));
      w.document.close();
    }
  }

  exportExcel() {
    const rows = Array.from({ length: this.quantity }, (_, i) => ({
      N: i + 1,
      Modo: this.mode,
      Codigo_o_texto: this.value,
      TipoBarras: this.mode === 'barcode' ? this.type : '',
      Descripcion: this.description,
      Precio: this.price,
      Lote: this.lot,
      Encabezado: this.headerText,
      Pie: this.footerText,
      Ancho_mm: this.labelWidthMm,
      Alto_mm: this.labelHeightMm,
      MostrarValor: this.showValue,
      TieneLogo: !!this.logoDataUrl,
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Etiquetas');
    XLSX.writeFile(wb, 'etiquetas.xlsx');
  }

  async exportPng() {
    const cfg = (this.useBatch && this.selectedBatchIndex != null && this.batchItems[this.selectedBatchIndex])
      ? this.batchItems[this.selectedBatchIndex]
      : this.getCurrentConfig();
    const dataUrl = await this.buildLabelDataUrlFrom(cfg, this.dpi);
    const safe = (cfg.value || 'etiqueta').replace(/[^a-z0-9_-]+/gi, '_').slice(0, 40);
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = `etiqueta_${safe}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  }

  async exportPngPorFilaZip() {
    if (!this.useBatch || !this.batchItems.length) return;
    const zip = new JSZip();

    // CSV resumen
    const rows: Array<Record<string, any>> = [];

    for (let i = 0; i < this.batchItems.length; i++) {
      if (this.batchErrors[i]) continue;
      const cfg = this.batchItems[i];
      const count = Math.max(1, Math.floor(cfg.qty || 1));

      const safe = (cfg.value || `fila-${i+1}`).replace(/[^a-z0-9_-]+/gi, '_').slice(0, 40);
      const folder = zip.folder(`etiquetas/${safe}`) ?? zip;

      for (let k = 0; k < count; k++) {
        const dataUrl = await this.buildLabelDataUrlFrom(cfg, this.dpi);
        const bytes = this.dataURLtoUint8Array(dataUrl);
        const name = count > 1 ? `etiqueta_${safe}_${k+1}.png` : `etiqueta_${safe}.png`;
        folder.file(name, bytes);
      }

      rows.push({
        index: i + 1,
        value: cfg.value,
        type: cfg.type,
        mode: cfg.mode,
        qty: count,
        description: cfg.description,
        price: cfg.price,
        lot: cfg.lot,
        width_mm: cfg.labelWidthMm,
        height_mm: cfg.labelHeightMm,
        margin_mm: cfg.marginMm,
        dpi: cfg.dpi,
        layout: cfg.layout,
        align: cfg.bodyAlign,
        header: cfg.headerText,
        footer: cfg.footerText,
        has_logo: !!cfg.logoDataUrl,
        folder: `etiquetas/${safe}`
      });
    }

    // Añadir CSV al zip
    const headers = [
      'index','value','type','mode','qty','description','price','lot','width_mm','height_mm','margin_mm','dpi','layout','align','header','footer','has_logo','folder'
    ];
    const csv = this.toCsv(headers, rows);
    zip.file('resumen.csv', csv);

    const blob = await zip.generateAsync({ type: 'blob' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'etiquetas_png.zip';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  private toCsv(headers: string[], data: Array<Record<string, any>>): string {
    const esc = (v: any) => {
      const s = v == null ? '' : String(v);
      // Escapar comillas dobles
      const e = s.replace(/"/g, '""');
      // Envolver en comillas si tiene coma, comillas o salto de línea
      return /[",\n]/.test(e) ? `"${e}"` : e;
    };
    const lines: string[] = [];
    lines.push(headers.join(','));
    for (const row of data) {
      lines.push(headers.map(h => esc(row[h])).join(','));
    }
    return lines.join('\n');
  }

  private dataURLtoUint8Array(dataUrl: string): Uint8Array {
    const base64 = dataUrl.split(',')[1] || '';
    const bin = atob(base64);
    const len = bin.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  onLogoSelected(evt: Event) {
    const input = evt.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      this.logoDataUrl = String(reader.result || '');
      this.renderPreview();
    };
    reader.readAsDataURL(file);
  }

  clearLogo() {
    this.logoDataUrl = undefined;
    this.renderPreview();
  }

  async onExcelSelected(evt: Event) {
    const input = evt.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    this.excelRows = XLSX.utils.sheet_to_json<any>(sheet, { defval: '' });
    this.batchItems = this.excelRows.map(r => this.mapRowToConfig(r)).filter(Boolean) as LabelConfig[];
    this.computeBatchValidation();
    this.useBatch = this.batchItems.length > 0;
    this.selectedBatchIndex = this.batchItems.length ? 0 : null;
  }

  private mapRowToConfig(r: any): LabelConfig | null {
    const alias = (canonical: string) => {
      const a = this.customMap[canonical];
      return a && a.trim() ? [a.trim()] : [];
    };
    const get = (keys: string[], def: any = '') => {
      // Anteponer alias de mapeo personalizado
      const withAlias = [...alias(keys[0] || ''), ...keys];
      for (const k of withAlias) {
        if (k && r[k] !== undefined && r[k] !== null && r[k] !== '') return r[k];
      }
      return def;
    };
    const num = (v: any, d = 0) => {
      const n = Number(v);
      return isFinite(n) ? n : d;
    };
    const bool = (v: any, d = false) => {
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      if (typeof v === 'string') return ['true','1','si','sí','y','yes'].includes(v.toLowerCase());
      return d;
    };
    const mode = String(get(['Modo','mode'], this.mode)).toLowerCase() === 'qr' ? 'qr' : 'barcode';
    const alignMap: Record<string, Align> = { left: 'left', izquierda: 'left', center:'center', centro:'center', right:'right', derecha:'right' };
    const layoutMap: Record<string, LabelLayout> = { classic:'classic', clasico:'classic', clásico:'classic', logoleft:'logoLeft', 'logo izquierda':'logoLeft', codetop:'codeTop', 'codigo arriba':'codeTop', 'código arriba':'codeTop' };

    const cfg: LabelConfig = {
      value: String(get(['Codigo','Código','Valor','Value','Codigo_o_texto'], this.value)),
      type: (String(get(['Tipo','Type'], this.type)) as BarcodeType) || 'CODE128',
      mode,
      showValue: bool(get(['MostrarValor','ShowValue'], this.showValue), this.showValue),
      description: String(get(['Descripcion','Descripción','Description'], '')),
      price: String(get(['Precio','Price'], '')),
      lot: String(get(['Lote','Lot'], '')),
      headerText: String(get(['Encabezado','Header'], '')),
      footerText: String(get(['Pie','Footer'], '')),
      logoDataUrl: String(get(['Logo','LogoUrl'], '')) || undefined,
      labelWidthMm: num(get(['Ancho_mm','Width_mm'], this.labelWidthMm), this.labelWidthMm),
      labelHeightMm: num(get(['Alto_mm','Height_mm'], this.labelHeightMm), this.labelHeightMm),
      marginMm: num(get(['Margen_mm','Margin_mm'], this.marginMm), this.marginMm),
      dpi: num(get(['DPI','dpi'], this.dpi), this.dpi),
      layout: layoutMap[String(get(['Layout'], this.layout)).toLowerCase()] || this.layout,
      bodyAlign: alignMap[String(get(['Alineacion','Alineación','Align'], this.bodyAlign)).toLowerCase()] || this.bodyAlign,
      headerScale: num(get(['EscalaEncabezado','HeaderScale'], this.headerScale), this.headerScale),
      bodyScale: num(get(['EscalaCuerpo','BodyScale'], this.bodyScale), this.bodyScale),
      footerScale: num(get(['EscalaPie','FooterScale'], this.footerScale), this.footerScale),
      qty: num(get(['Cantidad','Qty','Quantity'], 1), 1),
    };
    return cfg;
  }

  private computeBatchValidation() {
    this.batchErrors = this.batchItems.map((cfg, idx) => this.validateConfig(cfg));
  }

  applyCustomMapping() {
    if (!this.excelRows.length) return;
    this.batchItems = this.excelRows.map(r => this.mapRowToConfig(r)).filter(Boolean) as LabelConfig[];
    this.computeBatchValidation();
    if (this.batchItems.length) this.selectedBatchIndex = Math.min(this.selectedBatchIndex ?? 0, this.batchItems.length - 1);
    this.renderPreview();
  }

  setSelectedBatchIndex(i: string | number) {
    const idx = typeof i === 'string' ? Number(i) : i;
    if (Number.isFinite(idx) && idx >= 0 && idx < this.batchItems.length) {
      this.selectedBatchIndex = idx;
      this.renderPreview();
    }
  }

  private validateConfig(c: LabelConfig): string {
    // Reglas básicas por tipo
    const onlyDigits = (s: string) => /^[0-9]+$/.test(s);
    const len = c.value.length;
    if (!c.value) return 'Valor vacío';
    if (c.mode === 'barcode') {
      switch (c.type) {
        case 'EAN13':
          if (!onlyDigits(c.value) || !(len === 12 || len === 13)) return 'EAN13 debe ser 12 o 13 dígitos';
          break;
        case 'EAN8':
          if (!onlyDigits(c.value) || !(len === 7 || len === 8)) return 'EAN8 debe ser 7 u 8 dígitos';
          break;
        case 'UPC':
          if (!onlyDigits(c.value) || !(len === 11 || len === 12)) return 'UPC-A debe ser 11 o 12 dígitos';
          break;
        case 'ITF14':
          if (!onlyDigits(c.value) || !(len === 13 || len === 14)) return 'ITF-14 debe ser 13 o 14 dígitos';
          break;
        case 'MSI':
        case 'pharmacode':
        case 'codabar':
        case 'CODE128':
        default:
          // sin validación estricta
          break;
      }
    }
    if (c.labelWidthMm <= 0 || c.labelHeightMm <= 0) return 'Tamaño etiqueta inválido';
    if (c.dpi < 72) return 'DPI muy bajo (<72)';
    return '';
  }

  async exportPdfPorFila() {
    // Genera un PDF por cada fila válida en batch
    if (!this.useBatch || !this.batchItems.length) return;
    const page = this.getPageSizeMm(this.pageFormat);
    const gap = this.marginMm;
    const labelW = this.labelWidthMm;
    const labelH = this.labelHeightMm;
    const cols = Math.max(1, Math.floor((page.w - gap) / (labelW + gap)));
    const rows = Math.max(1, Math.floor((page.h - gap) / (labelH + gap)));
    const perPage = cols * rows;

    for (let r = 0; r < this.batchItems.length; r++) {
      if (this.batchErrors[r]) continue; // salta inválidos
      const cfg = this.batchItems[r];
      const count = Math.max(1, Math.floor(cfg.qty || 1));
      const doc = new jsPDF({ unit: 'mm', format: this.pageFormat, compress: true });
      for (let i = 0; i < count; i++) {
        const idx = i % perPage;
        if (i > 0 && idx === 0) doc.addPage(this.pageFormat);
        const col = idx % cols;
        const row = Math.floor(idx / cols);
        const x = gap + col * (labelW + gap);
        const y = gap + row * (labelH + gap);
        const imgData = await this.buildLabelDataUrlFrom(cfg, this.dpi);
        doc.addImage(imgData, 'PNG', x, y, labelW, labelH);
      }
      const safe = (cfg.value || `fila-${r+1}`).replace(/[^a-z0-9_-]+/gi, '_').slice(0, 40);
      doc.save(`etiquetas_${safe}.pdf`);
    }
  }

  private getCurrentConfig(): LabelConfig {
    return {
      value: this.value,
      type: this.type,
      mode: this.mode,
      showValue: this.showValue,
      description: this.description,
      price: this.price,
      lot: this.lot,
      headerText: this.headerText,
      footerText: this.footerText,
      logoDataUrl: this.logoDataUrl,
      labelWidthMm: this.labelWidthMm,
      labelHeightMm: this.labelHeightMm,
      marginMm: this.marginMm,
      dpi: this.dpi,
      layout: this.layout,
      bodyAlign: this.bodyAlign,
      headerScale: this.headerScale,
      bodyScale: this.bodyScale,
      footerScale: this.footerScale,
      qty: this.quantity,
    };
  }

  private updateFormError() {
    if (this.useBatch) {
      this.formError = null;
      return;
    }
    const err = this.validateConfig(this.getCurrentConfig());
    this.formError = err || null;
  }

  resetForm() {
    // Confirmaciones
    if (this.confirmOnResetForm) {
      if (this.batchItems.length || this.useBatch) {
        const ok = confirm('Esto limpiará el formulario y también los datos cargados (Excel/lote). ¿Continuar?');
        if (!ok) return;
      } else {
        const ok = confirm('Esto restaurará el formulario a sus valores por defecto. ¿Continuar?');
        if (!ok) return;
      }
    }

    // Valores por defecto
    this.value = '123456789012';
    this.type = 'CODE128';
    this.mode = 'barcode';
    this.showValue = true;

    this.description = '';
    this.price = '';
    this.lot = '';
    this.headerText = '';
    this.footerText = '';
    this.logoDataUrl = undefined;

    this.selectedTemplate = '50 x 30 mm';
    this.labelWidthMm = 50;
    this.labelHeightMm = 30;
    this.quantity = 1;
    this.pageFormat = 'a4';
    this.marginMm = 5;
    this.dpi = 300;

    this.layout = 'classic';
    this.bodyAlign = 'center';
    this.headerScale = 0.10;
    this.bodyScale = 0.08;
    this.footerScale = 0.07;

    // Limpiar Excel y lote
    this.batchItems = [];
    this.batchErrors = [];
    this.excelRows = [];
    this.selectedBatchIndex = null;
    this.useBatch = false;

    // Mantener mapeo personalizado sin cambios

    this.renderPreview();
  }

  onToggleConfirmPrefs() {
    try {
      localStorage.setItem('confirmOnResetForm', String(this.confirmOnResetForm));
      localStorage.setItem('confirmOnResetMapping', String(this.confirmOnResetMapping));
    } catch {}
  }

  resetMapping() {
    if (this.confirmOnResetMapping) {
      const ok = confirm('Esto limpiará el mapeo de columnas personalizado. ¿Continuar?');
      if (!ok) return;
    }
    this.customMap = {
      Codigo: '', Tipo: '', Modo: '', Cantidad: '', Descripcion: '', Precio: '', Lote: '',
      Encabezado: '', Pie: '', Ancho_mm: '', Alto_mm: '', Margen_mm: '', DPI: '', Layout: '',
      Alineacion: '', EscalaEncabezado: '', EscalaCuerpo: '', EscalaPie: '', LogoUrl: '',
    };
  }

  resetBatch() {
    if (this.confirmOnResetForm && (this.batchItems.length || this.useBatch)) {
      const ok = confirm('Esto limpiará los datos cargados del Excel/lote (no cambia el formulario ni el mapeo). ¿Continuar?');
      if (!ok) return;
    }
    this.batchItems = [];
    this.batchErrors = [];
    this.excelRows = [];
    this.selectedBatchIndex = null;
    this.useBatch = false;
  }

  downloadExcelTemplate() {
    const headers = {
      Codigo: '', Tipo: 'CODE128', Modo: 'barcode', Cantidad: 1,
      Descripcion: '', Precio: '', Lote: '', Encabezado: '', Pie: '',
      Ancho_mm: this.labelWidthMm, Alto_mm: this.labelHeightMm, Margen_mm: this.marginMm, DPI: this.dpi,
      Layout: this.layout, Alineacion: this.bodyAlign, EscalaEncabezado: this.headerScale, EscalaCuerpo: this.bodyScale, EscalaPie: this.footerScale,
      LogoUrl: ''
    };
    const sample = [
      { ...headers, Codigo: '7501234567890', Tipo: 'EAN13', Cantidad: 5, Descripcion: 'Producto A', Precio: '10.99' },
      { ...headers, Codigo: 'ABC-001', Tipo: 'CODE128', Cantidad: 2, Descripcion: 'Producto B', Precio: '7.50' },
    ];
    const ws = XLSX.utils.json_to_sheet(sample, { header: Object.keys(headers) });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Plantilla');
    XLSX.writeFile(wb, 'plantilla_etiquetas.xlsx');
  }

  private mmToPx(mm: number, dpi: number) {
    return (mm / 25.4) * dpi;
  }

  private async buildLabelDataUrl(dpi: number): Promise<string> {
    return this.buildLabelDataUrlFrom(this.getCurrentConfig(), dpi);
  }

  private async buildLabelDataUrlFrom(cfg: LabelConfig, dpi: number): Promise<string> {
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(10, Math.round(this.mmToPx(cfg.labelWidthMm, dpi)));
    canvas.height = Math.max(10, Math.round(this.mmToPx(cfg.labelHeightMm, dpi)));
    await this.drawLabelToCanvasWith(cfg, canvas, dpi);
    return canvas.toDataURL('image/png');
  }

  private async drawLabelToCanvas(canvas: HTMLCanvasElement, dpi: number) {
    await this.drawLabelToCanvasWith(this.getCurrentConfig(), canvas, dpi);
  }

  private async drawLabelToCanvasWith(c: LabelConfig, canvas: HTMLCanvasElement, dpi: number) {
    const ctx = canvas.getContext('2d')!;
    const W = canvas.width;
    const H = canvas.height;
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, W, H);
    ctx.fillStyle = '#000';
    ctx.textBaseline = 'top';

    // Margen interno dentro de la etiqueta (3% del alto)
    const pad = Math.round(H * 0.03);
    let y = pad;

    // Encabezado opcional
    if (c.headerText) {
      ctx.font = `${Math.max(10, Math.round(H * c.headerScale))}px sans-serif`;
      const text = c.headerText;
      const metrics = ctx.measureText(text);
      const tx = Math.max(pad, Math.min(W - pad - metrics.width, Math.round((W - metrics.width) / 2)));
      ctx.fillText(text, tx, y);
      y += Math.round(H * 0.12);
    }

    // Logo opcional (arriba-izquierda)
    const logoBox = { w: Math.round(W * 0.22), h: Math.round(H * 0.22) };
    if (c.logoDataUrl) {
      try {
        const img = await this.loadImage(c.logoDataUrl!);
        const scale = Math.min(logoBox.w / img.width, logoBox.h / img.height, 1);
        const lw = Math.round(img.width * scale);
        const lh = Math.round(img.height * scale);
        ctx.drawImage(img, pad, y, lw, lh);
      } catch {}
    }

    // Área disponible y layout
    let areaY = y;
    let areaHeight = Math.round(H * 0.55);

    if (c.layout === 'logoLeft' && c.logoDataUrl) {
      // Reservar una franja izquierda para el logo (ancho 22% W)
      try {
        const img = await this.loadImage(c.logoDataUrl!);
        const lw = Math.round(W * 0.22);
        const lh = Math.round(H * 0.5);
        const scale = Math.min(lw / img.width, lh / img.height, 1);
        const drawW = Math.round(img.width * scale);
        const drawH = Math.round(img.height * scale);
        ctx.drawImage(img, pad, y, drawW, drawH);
      } catch {}
    }

    if (c.layout === 'codeTop') {
      areaY = y;
      areaHeight = Math.round(H * 0.45);
    }

    if (c.mode === 'barcode') {
      // Render a un canvas temporal para el código de barras y pegarlo
      const bcCanvas = document.createElement('canvas');
      bcCanvas.width = Math.max(10, Math.round(W - pad * 2));
      bcCanvas.height = Math.max(10, Math.round(areaHeight));
      try {
        JsBarcode(bcCanvas, c.value, {
          format: c.type,
          lineColor: '#000',
          background: '#ffffff',
          displayValue: c.showValue,
          margin: Math.max(4, Math.round(H * 0.01)),
          width: Math.max(1, Math.round(W / 200)),
          height: Math.max(10, Math.round(areaHeight * 0.75)),
          fontSize: Math.max(8, Math.round(H * c.bodyScale)),
        } as any);
        this.renderError = null;
      } catch (e: any) {
        this.renderError = `No se pudo generar el código de barras: ${e?.message || e}`;
        // Dibujar placeholder de error
        ctx.fillStyle = '#fee';
        ctx.fillRect(pad, areaY, W - pad * 2, areaHeight);
        ctx.fillStyle = '#b00020';
        ctx.font = `${Math.max(10, Math.round(H * 0.07))}px sans-serif`;
        ctx.fillText('Error al generar código', pad + 6, areaY + 6);
        return;
      }
      // Centrar horizontalmente
      let dx = Math.round((W - bcCanvas.width) / 2);
      if (c.layout === 'logoLeft' && c.logoDataUrl) {
        // desplazar a la derecha dejando margen para el logo
        dx = Math.max(dx, Math.round(W * 0.25));
      }
      ctx.drawImage(bcCanvas, dx, areaY);
    } else {
      // QR code centrado
      let qrSize = Math.round(Math.min(W, H) * 0.6);
      if (c.layout === 'logoLeft' && c.logoDataUrl) {
        qrSize = Math.round(qrSize * 0.75);
      }
      try {
        const qrDataUrl = await QRCode.toDataURL(c.value, { width: qrSize, margin: 1 });
        const img = await this.loadImage(qrDataUrl);
        const x = Math.round((W - qrSize) / 2);
        const yq = areaY + Math.round((areaHeight - qrSize) / 2);
        ctx.drawImage(img, x, yq, qrSize, qrSize);
        this.renderError = null;
      } catch (e: any) {
        this.renderError = `No se pudo generar el QR: ${e?.message || e}`;
        ctx.fillStyle = '#fee';
        ctx.fillRect(pad, areaY, W - pad * 2, areaHeight);
        ctx.fillStyle = '#b00020';
        ctx.font = `${Math.max(10, Math.round(H * 0.07))}px sans-serif`;
        ctx.fillText('Error al generar QR', pad + 6, areaY + 6);
        return;
      }
    }

    // Texto descriptivo debajo
    let textY = areaY + areaHeight + Math.round(H * 0.02);
    ctx.font = `${Math.max(8, Math.round(H * c.bodyScale))}px sans-serif`;
    const lineGap = Math.max(2, Math.round(H * 0.01));

    const lines: string[] = [];
    if (c.description) lines.push(c.description);
    if (c.price) lines.push(`Precio: ${c.price}`);
    if (c.lot) lines.push(`Lote: ${c.lot}`);

    for (const line of lines) {
      this.drawAlignedText(ctx, line, W, textY, c.bodyAlign);
      textY += Math.round(H * c.bodyScale * 1.2) + lineGap;
    }

    // Pie de página opcional
    if (c.footerText) {
      ctx.font = `${Math.max(8, Math.round(H * c.footerScale))}px sans-serif`;
      this.drawAlignedText(ctx, c.footerText, W, H - pad - Math.round(H * c.footerScale), c.bodyAlign);
    }
  }

  private drawCenteredText(ctx: CanvasRenderingContext2D, text: string, width: number, y: number) {
    const m = ctx.measureText(text);
    const x = Math.round((width - m.width) / 2);
    ctx.fillText(text, x, y);
  }

  private drawAlignedText(ctx: CanvasRenderingContext2D, text: string, width: number, y: number, align: Align) {
    const m = ctx.measureText(text);
    let x = 0;
    if (align === 'center') x = Math.round((width - m.width) / 2);
    else if (align === 'right') x = Math.round(width - m.width - 4);
    else x = 4;
    ctx.fillText(text, x, y);
  }

  private loadImage(src: string): Promise<HTMLImageElement> {
    return new Promise((resolve, reject) => {
      const img = new Image();
      img.onload = () => resolve(img);
      img.onerror = reject;
      img.src = src;
    });
  }

  private getPageSizeMm(format: 'a4' | 'letter') {
    if (format === 'a4') return { w: 210, h: 297 };
    return { w: 216, h: 279 };
  }
}
