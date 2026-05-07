/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import { 
  FileSpreadsheet, 
  Plus, 
  Trash2, 
  Copy, 
  Check, 
  MessageSquare, 
  Download,
  Search,
  Settings2,
  Trash,
  ClipboardPaste,
  X,
  Bell
} from 'lucide-react';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';
import { cn, generateMessage } from './lib/utils';

interface DataRow {
  id: string;
  nombreNino: string;
  nombreAdulto: string; // Referencial
  telefono: string;
  rutAdulto: string;
  tipoIngreso: string;
  rit: string;
}

const DEFAULT_ROW: DataRow = {
  id: '',
  nombreNino: '',
  nombreAdulto: '',
  telefono: '',
  rutAdulto: '',
  tipoIngreso: 'Tribunal',
  rit: '',
};

const FIXED_HEADERS = [
  "NOMBRES NNA",
  "PRIMER APELLIDO",
  "SEGUNDO APELLIDO",
  "RUT",
  "FECHA ENTREGA CAUSA",
  "RIT",
  "OFICIO DE INGRESO",
  "AUDIENCIA",
  "ADULTA/O RESPONSABLE",
  "RELACIÓN CON NNA",
  "RUT DE P/M",
  "TELÉFONO",
  "DIRECCIÓN"
];

export default function App() {
  const [data, setData] = useState<DataRow[]>([]);
  const [mes, setMes] = useState('Mayo 2026');
  const [remitente, setRemitente] = useState('Arles García, secretario');
  const [copiedId, setCopiedId] = useState<string | null>(null);
  const [showPasteModal, setShowPasteModal] = useState(false);
  const [showSettingsModal, setShowSettingsModal] = useState(false);
  const [showBulkConfirmModal, setShowBulkConfirmModal] = useState(false);
  const [showClearConfirmModal, setShowClearConfirmModal] = useState(false);
  const [searchTerm, setSearchTerm] = useState('');
  const [pasteText, setPasteText] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processImportedData = (rawData: any[]) => {
    const grouped = new Map<string, DataRow>();

    rawData.forEach((row) => {
      const nombresNna = row['NOMBRES NNA'] || '';
      const primerApellido = row['PRIMER APELLIDO'] || '';
      const segundoApellido = row['SEGUNDO APELLIDO'] || '';
      const fullNino = `${nombresNna} ${primerApellido} ${segundoApellido}`.trim() || row['Nombre del niño'] || row['Nino'] || row['Nombre Niño'] || row['nombreNino'] || '';

      const rit = String(row['RIT'] || row['rit'] || row['Radicado'] || row['Causa'] || row['radicado'] || row['causa'] || '').trim();
      const telefono = String(row['TELÉFONO'] || row['Telefono'] || row['telefono'] || row['Teléfono'] || '').trim();
      
      // Grouping key: Use RIT if available, else fallback to phone
      const groupKey = rit !== '' ? `rit:${rit}` : `tel:${telefono}`;

      if (grouped.has(groupKey)) {
        const existing = grouped.get(groupKey)!;
        // Only append if name is not already present
        if (!existing.nombreNino.split(', ').includes(fullNino)) {
          existing.nombreNino += `, ${fullNino}`;
        }
      } else {
        grouped.set(groupKey, {
          id: Math.random().toString(36).substr(2, 9),
          nombreNino: fullNino,
          nombreAdulto: row['ADULTA/O RESPONSABLE'] || row['Adulto'] || row['nombreAdulto'] || row['Nombre del adulto responsable'] || '',
          telefono: telefono,
          rutAdulto: row['RUT DE P/M'] || row['RUT'] || row['rutAdulto'] || row['RUT del adulto responsable'] || '',
          tipoIngreso: row['Tipo de ingreso'] || row['tipoIngreso'] || (rit ? 'Tribunal' : 'OLN'),
          rit: rit,
        });
      }
    });

    return Array.from(grouped.values());
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target?.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const rawData = XLSX.utils.sheet_to_json(ws) as any[];
      
      const formattedData = processImportedData(rawData);
      setData(prev => [...prev, ...formattedData]);
    };
    reader.readAsBinaryString(file);
  };

  const handlePaste = () => {
    if (!pasteText.trim()) return;

    const lines = pasteText.trim().split('\n');
    let header: string[] = [];
    let dataLines: string[] = [];

    // Check if the first line is the header row or data
    const firstLine = lines[0].toLowerCase();
    const isFirstLineHeader = firstLine.includes('nombre') || firstLine.includes('nna') || firstLine.includes('rut');

    if (isFirstLineHeader) {
      header = lines[0].split('\t').map(h => h.trim());
      dataLines = lines.slice(1);
    } else {
      // Use fixed headers and assume all lines are data
      header = FIXED_HEADERS;
      dataLines = lines;
    }
    
    const rows = dataLines.map(line => {
      const cols = line.split('\t');
      const obj: any = {};
      header.forEach((h, i) => {
        obj[h] = cols[i]?.trim() || '';
      });
      return obj;
    });

    const formattedData = processImportedData(rows);
    setData(prev => [...prev, ...formattedData]);
    setPasteText('');
    setShowPasteModal(false);
  };

  const handleBulkCopy = () => {
    const allMessages = data.map(d => {
      const allNinosForThisPhone = data
        .filter(row => row.telefono === d.telefono)
        .map(row => row.nombreNino);
      return generateMessage({ ...d, ninos: allNinosForThisPhone, mes, remitente });
    }).join('\n\n---\n\n');
    
    copyToClipboard(allMessages, 'bulk');
    setShowBulkConfirmModal(false);
    alert('Todos los mensajes han sido copiados al portapapeles.');
  };

  const handleBulkSend = () => {
    // Group all data by phone number to send ONE single message per contact
    const phoneGroups = new Map<string, { ninos: Set<string>, rits: Set<string>, originalRow: DataRow }>();
    
    data.forEach(row => {
      const phone = row.telefono.trim();
      if (!phoneGroups.has(phone)) {
        phoneGroups.set(phone, { 
          ninos: new Set(row.nombreNino.split(',').map(n => n.trim())), 
          rits: new Set(row.rit ? [row.rit.trim()] : []),
          originalRow: row 
        });
      } else {
        const group = phoneGroups.get(phone)!;
        row.nombreNino.split(',').forEach(n => group.ninos.add(n.trim()));
        if (row.rit) group.rits.add(row.rit.trim());
      }
    });

    const uniqueContacts = Array.from(phoneGroups.values());
    
    if (uniqueContacts.length > 5) {
      if (!confirm(`Se intentarán abrir ${uniqueContacts.length} pestañas de WhatsApp. ¿Deseas continuar? (Asegúrate de permitir ventanas emergentes/pop-ups)`)) {
        return;
      }
    }

    uniqueContacts.forEach((contact, index) => {
      setTimeout(() => {
        const { ninos, rits, originalRow } = contact;
        const msg = generateMessage({ 
          ...originalRow, 
          ninos: Array.from(ninos), 
          rit: Array.from(rits).join(' / '),
          mes, 
          remitente 
        });
        
        const phoneClean = originalRow.telefono.replace(/\s/g, '');
        const phoneWithCountry = phoneClean.length === 9 ? '56' + phoneClean : phoneClean.replace(/\+/g, '');
        const url = `https://wa.me/${phoneWithCountry}?text=${encodeURIComponent(msg)}`;
        window.open(url, '_blank');
      }, index * 1200); // Slightly longer delay for better reliability
    });
    
    setShowBulkConfirmModal(false);
  };

  const addRow = () => {
    setData([
      ...data,
      { ...DEFAULT_ROW, id: Math.random().toString(36).substr(2, 9) }
    ]);
  };

  const updateRow = (id: string, field: keyof DataRow, value: string) => {
    setData(data.map(row => row.id === id ? { ...row, [field]: value } : row));
  };

  const removeRow = (id: string) => {
    setData(data.filter(row => row.id !== id));
  };

  const handleClearAll = () => {
    setData([]);
    setShowClearConfirmModal(false);
  };

  const copyToClipboard = (text: string, id: string) => {
    navigator.clipboard.writeText(text);
    setCopiedId(id);
    setTimeout(() => setCopiedId(null), 2000);
  };

  const messagesReady = data.length;
  const tribunalCases = data.filter(d => d.tipoIngreso === 'Tribunal').length;

  return (
    <div className="min-h-screen bg-[#FDFCFB] text-[#1A1A18] font-sans pb-24">
      {/* Premium Navigation Header */}
      <nav className="sticky top-0 z-50 bg-white/80 backdrop-blur-xl border-b border-[#E5E0DA] px-4 py-3 md:px-8">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-[#5a5a40] rounded-xl flex items-center justify-center text-white shadow-lg shadow-[#5a5a40]/20">
              <MessageSquare size={20} />
            </div>
            <div>
              <h1 className="text-lg font-bold tracking-tight text-[#1A1A18]">DCE San Bernardo 2</h1>
              <div className="flex items-center gap-1.5 text-[10px] text-[#86867A] font-semibold uppercase tracking-wider">
                <span className="w-1.5 h-1.5 bg-emerald-500 rounded-full animate-pulse" />
                Sistema de Gestión
              </div>
            </div>
          </div>
          
          <div className="flex items-center gap-2">
            <button 
              onClick={() => setShowSettingsModal(true)}
              className="p-2 hover:bg-[#F7F5F2] rounded-full transition-all text-[#5a5a40]"
            >
              <Settings2 size={20} />
            </button>
            <div className="h-6 w-px bg-[#E5E0DA] mx-1" />
            <button 
              onClick={() => fileInputRef.current?.click()}
              className="hidden md:flex items-center gap-2 bg-[#1A1A18] text-white px-5 py-2.5 rounded-full text-xs font-bold hover:bg-black transition-all shadow-md shadow-black/10"
            >
              <Download size={14} />
              Importar
            </button>
            <button 
              onClick={() => setShowPasteModal(true)}
              className="p-2 bg-[#F7F5F2] text-[#5a5a40] rounded-full md:hidden"
            >
              <ClipboardPaste size={20} />
            </button>
          </div>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 py-8 md:px-8 space-y-10">
        {/* Modern Stats Grid */}
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-4 md:gap-6">
          {[
            { label: 'Mensajes Listos', value: new Set(data.map(d => d.telefono)).size, icon: MessageSquare, color: 'text-emerald-600', bg: 'bg-emerald-50' },
            { label: 'Total Registros', value: data.length, icon: FileSpreadsheet, color: 'text-[#5a5a40]', bg: 'bg-[#F7F5F2]' },
            { label: 'Tribunal', value: tribunalCases, icon: Bell, color: 'text-amber-600', bg: 'bg-amber-50' },
            { label: 'Estado', value: 'Activo', icon: Check, color: 'text-blue-600', bg: 'bg-blue-50' }
          ].map((stat, i) => (
            <motion.div 
              key={i}
              initial={{ opacity: 0, y: 20 }}
              animate={{ opacity: 1, y: 0 }}
              transition={{ delay: i * 0.1 }}
              className="bg-white p-5 md:p-6 rounded-[2rem] border border-[#E5E0DA] shadow-sm hover:shadow-md transition-all group"
            >
              <div className="flex items-center justify-between mb-4">
                <div className={cn("p-2.5 rounded-2xl transition-colors", stat.bg)}>
                  <stat.icon size={20} className={stat.color} />
                </div>
              </div>
              <p className="text-3xl font-bold tracking-tight text-[#1A1A18]">
                {typeof stat.value === 'number' ? stat.value.toString().padStart(2, '0') : stat.value}
              </p>
              <p className="text-[11px] font-bold text-[#86867A] uppercase tracking-wider mt-1">{stat.label}</p>
            </motion.div>
          ))}
        </div>

        {/* Search and Quick Actions */}
        <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
          <div className="relative flex-1 max-w-md group">
            <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-[#86867A] group-focus-within:text-[#5a5a40] transition-colors" size={18} />
            <input 
              type="text"
              placeholder="Buscar por niño, RIT o teléfono..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-12 pr-4 py-3.5 bg-white border border-[#E5E0DA] rounded-2xl text-sm focus:ring-4 focus:ring-[#5a5a40]/5 focus:border-[#5a5a40] outline-none transition-all placeholder:text-[#86867A]"
            />
          </div>

          <div className="flex items-center gap-3 overflow-x-auto pb-2 md:pb-0 no-scrollbar">
            <button 
              onClick={addRow}
              className="whitespace-nowrap px-6 py-3.5 bg-[#F7F5F2] text-[#5a5a40] rounded-2xl text-sm font-bold hover:bg-[#E5E0DA] transition-all flex items-center gap-2"
            >
              <Plus size={18} />
              Agregar Registro
            </button>
            {data.length > 0 && (
              <button 
                onClick={() => setShowClearConfirmModal(true)}
                className="whitespace-nowrap px-6 py-3.5 bg-rose-50 text-rose-600 rounded-2xl text-sm font-bold hover:bg-rose-100 transition-all flex items-center gap-2"
              >
                <Trash2 size={18} />
                Limpiar Todo
              </button>
            )}
          </div>
        </div>

        {/* Modern Data List (Replaces Table) */}
        <div className="space-y-4">
          <div className="flex items-center justify-between px-2">
            <h2 className="text-sm font-bold uppercase tracking-widest text-[#86867A]">Listado de Registros</h2>
            <div className="flex items-center gap-2 text-[10px] text-[#86867A] font-bold uppercase italic">
              <span>{data.length} En total</span>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            <AnimatePresence mode="popLayout">
              {data.length === 0 ? (
                <motion.div 
                  initial={{ opacity: 0 }}
                  animate={{ opacity: 1 }}
                  className="col-span-full py-20 text-center bg-white rounded-3xl border border-dashed border-[#E5E0DA]"
                >
                  <div className="w-16 h-16 bg-[#F7F5F2] rounded-full flex items-center justify-center mx-auto mb-4 text-[#5a5a40]">
                    <FileSpreadsheet size={32} />
                  </div>
                  <h3 className="text-lg font-bold text-[#1A1A18]">Sin datos cargados</h3>
                  <p className="text-sm text-[#86867A] mt-1">Sube un Excel o pega una tabla para comenzar.</p>
                </motion.div>
              ) : (
                data
                  .filter(row => {
                    const search = searchTerm.toLowerCase();
                    return (
                      row.nombreNino.toLowerCase().includes(search) ||
                      row.rit.toLowerCase().includes(search) ||
                      row.telefono.includes(search) ||
                      row.nombreAdulto.toLowerCase().includes(search)
                    );
                  })
                  .map((row) => {
                    const neighbors = data.filter(d => d.id !== row.id);
                    const siblingsInCase = neighbors.filter(d => d.rit === row.rit && row.rit !== '');
                    const othersWithSamePhone = neighbors.filter(d => d.telefono === row.telefono);
                    const ninosList = row.nombreNino.split(',').map(n => n.trim()).filter(n => n !== '');
                    const generatedMsg = generateMessage({ ...row, ninos: ninosList, mes, remitente });

                    return (
                      <motion.div 
                        key={row.id}
                        layout
                        initial={{ opacity: 0, scale: 0.95 }}
                        animate={{ opacity: 1, scale: 1 }}
                        exit={{ opacity: 0, scale: 0.9 }}
                        className="bg-white rounded-3xl border border-[#E5E0DA] shadow-sm hover:shadow-xl transition-all overflow-hidden flex flex-col group"
                      >
                        {/* Card Header */}
                        <div className="p-5 border-b border-[#F7F5F2] bg-white flex justify-between items-start">
                          <div className="space-y-1">
                            <div className="flex items-center gap-2">
                              <span className={cn(
                                "px-2 py-0.5 rounded-full text-[9px] font-bold uppercase",
                                row.tipoIngreso === 'Tribunal' ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700"
                              )}>
                                {row.tipoIngreso}
                              </span>
                              {siblingsInCase.length > 0 && (
                                <span className="px-2 py-0.5 rounded-full bg-blue-100 text-blue-700 text-[9px] font-bold uppercase">Hermanos</span>
                              )}
                            </div>
                            <div className="flex items-center gap-1.5 text-xs font-bold text-[#1A1A18] font-mono">
                              <MessageSquare size={12} className="text-[#86867A]" />
                              {row.telefono}
                            </div>
                          </div>
                          <button 
                            onClick={() => removeRow(row.id)}
                            className="p-1.5 text-[#86867A] hover:bg-rose-50 hover:text-rose-600 rounded-lg transition-all"
                          >
                            <X size={16} />
                          </button>
                        </div>

                        {/* Card Content */}
                        <div className="p-5 space-y-5 flex-1">
                          <div className="space-y-2">
                            <label className="text-[10px] font-bold text-[#86867A] uppercase tracking-wider block">Listado de Niños</label>
                            <div className="space-y-1.5">
                              {ninosList.map((name, idx) => (
                                <div key={idx} className="flex items-center gap-2 p-1.5 bg-[#F7F5F2] rounded-xl group/item">
                                  <div className="w-1.5 h-1.5 bg-[#5a5a40] rounded-full shrink-0" />
                                  <input 
                                    type="text" 
                                    value={name}
                                    onChange={(e) => {
                                      const names = row.nombreNino.split(',');
                                      names[idx] = e.target.value;
                                      updateRow(row.id, 'nombreNino', names.join(','));
                                    }}
                                    className="w-full bg-transparent border-none p-0 text-xs font-bold focus:ring-0 text-[#1A1A18]"
                                  />
                                  {idx > 0 && (
                                    <button 
                                      onClick={() => {
                                        const names = row.nombreNino.split(',').filter((_, i) => i !== idx);
                                        updateRow(row.id, 'nombreNino', names.join(','));
                                      }}
                                      className="opacity-0 group-hover/item:opacity-100 p-1 text-rose-500 transition-opacity"
                                    >
                                      <Trash size={10} />
                                    </button>
                                  )}
                                </div>
                              ))}
                              <button 
                                onClick={() => updateRow(row.id, 'nombreNino', row.nombreNino + ', Nuevo Niño')}
                                className="w-full py-2 border border-dashed border-[#E5E0DA] rounded-xl text-[10px] font-bold text-[#86867A] hover:text-[#5a5a40] hover:bg-[#F7F5F2] transition-all flex items-center justify-center gap-1.5"
                              >
                                <Plus size={12} />
                                Añadir Niño
                              </button>
                            </div>
                          </div>

                          <div className="grid grid-cols-2 gap-4">
                            <div className="space-y-1.5">
                              <label className="text-[10px] font-bold text-[#86867A] uppercase tracking-wider">Causa (RIT)</label>
                              <input 
                                type="text"
                                value={row.rit}
                                onChange={(e) => updateRow(row.id, 'rit', e.target.value)}
                                className="w-full px-3 py-2 bg-[#F7F5F2] rounded-xl text-xs font-mono font-bold focus:ring-2 focus:ring-[#5a5a40]/10 border-none"
                                placeholder="P-000-2025"
                              />
                            </div>
                            <div className="space-y-1.5">
                              <label className="text-[10px] font-bold text-[#86867A] uppercase tracking-wider">RUT Adulto</label>
                              <input 
                                type="text"
                                value={row.rutAdulto}
                                onChange={(e) => updateRow(row.id, 'rutAdulto', e.target.value)}
                                className="w-full px-3 py-2 bg-[#F7F5F2] rounded-xl text-xs font-medium focus:ring-2 focus:ring-[#5a5a40]/10 border-none"
                                placeholder="12.345.678-9"
                              />
                            </div>
                          </div>

                          <div className="space-y-2">
                             <label className="text-[10px] font-bold text-[#86867A] uppercase tracking-wider flex justify-between">
                              Vista Previa
                              <span className="italic font-normal lowercase">{generatedMsg.length} caracteres</span>
                             </label>
                             <div className="bg-[#FAF9F7] p-4 rounded-2xl text-[11px] leading-relaxed text-[#5a5a40] max-h-32 overflow-y-auto no-scrollbar italic whitespace-pre-wrap">
                               {generatedMsg}
                             </div>
                          </div>
                        </div>

                        {/* Card Footer Actions */}
                        <div className="p-3 bg-[#FAF9F7] flex gap-2">
                          <button 
                            onClick={() => copyToClipboard(generatedMsg, row.id)}
                            className={cn(
                              "flex-1 py-3 rounded-2xl text-xs font-bold transition-all flex items-center justify-center gap-2",
                              copiedId === row.id ? "bg-emerald-600 text-white" : "bg-[#1A1A18] text-white hover:bg-black"
                            )}
                          >
                            {copiedId === row.id ? <Check size={14} /> : <Copy size={14} />}
                            {copiedId === row.id ? 'Copiado' : 'Copiar'}
                          </button>
                          <a 
                            href={`https://wa.me/${row.telefono.replace(/\s/g, '').length === 9 ? '56' + row.telefono.replace(/\s/g, '') : row.telefono.replace(/\+/g, '').replace(/\s/g, '')}?text=${encodeURIComponent(generatedMsg)}`}
                            target="_blank"
                            rel="noopener noreferrer"
                            className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-2xl text-xs font-bold transition-all flex items-center justify-center gap-2"
                          >
                            <MessageSquare size={14} />
                            WhatsApp
                          </a>
                        </div>
                      </motion.div>
                    );
                  })
              )}
            </AnimatePresence>
          </div>
        </div>
      </main>

      {/* Floating Action Bar (Mobile Premium) */}
      <AnimatePresence>
        {data.length > 0 && (
          <motion.div 
            initial={{ y: 100 }}
            animate={{ y: 0 }}
            exit={{ y: 100 }}
            className="fixed bottom-6 left-4 right-4 z-40 max-w-lg mx-auto"
          >
            <div className="bg-[#1A1A18] p-2 rounded-[2rem] shadow-2xl flex items-center justify-between border border-white/10 backdrop-blur-md">
              <div className="pl-6 flex flex-col items-start leading-none">
                <span className="text-[10px] font-bold text-white/50 uppercase tracking-widest">Listo para enviar</span>
                <span className="text-sm font-bold text-white">{new Set(data.map(d => d.telefono)).size} <span className="text-xs font-normal text-white/60">Contactos</span></span>
              </div>
              <button 
                onClick={() => setShowBulkConfirmModal(true)}
                className="bg-white text-black px-8 py-3.5 rounded-[1.5rem] font-bold text-sm hover:bg-[#F7F5F2] transition-all flex items-center gap-2"
              >
                Asistente de Envío
                <MessageSquare size={16} />
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
      {/* Settings Modal */}
      <AnimatePresence>
        {showSettingsModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowSettingsModal(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="relative bg-white w-full max-w-md rounded-[2.5rem] shadow-2xl p-8 md:p-10 flex flex-col gap-8"
            >
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 bg-[#F7F5F2] rounded-2xl flex items-center justify-center text-[#5a5a40]">
                    <Settings2 size={24} />
                  </div>
                  <div>
                    <h3 className="text-xl font-bold tracking-tight text-[#1A1A18]">Configuración</h3>
                    <p className="text-xs text-[#86867A] font-medium uppercase tracking-wider">Ajustes del Sistema</p>
                  </div>
                </div>
                <button 
                  onClick={() => setShowSettingsModal(false)}
                  className="p-2 hover:bg-[#F7F5F2] rounded-full transition-colors text-[#86867A]"
                >
                  <X size={20} />
                </button>
              </div>

              <div className="space-y-6">
                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-[#86867A] uppercase tracking-widest ml-1">Firma del Remitente</label>
                  <input 
                    type="text" 
                    value={remitente}
                    onChange={(e) => setRemitente(e.target.value)}
                    placeholder="Ej: Arles García"
                    className="w-full px-5 py-4 bg-[#F7F5F2] border-none rounded-2xl focus:ring-4 focus:ring-[#5a5a40]/5 focus:bg-white transition-all font-semibold outline-none text-[#1A1A18]"
                  />
                </div>

                <div className="space-y-2">
                  <label className="text-[10px] font-bold text-[#86867A] uppercase tracking-widest ml-1">Mes del Proceso</label>
                  <input 
                    type="text" 
                    value={mes}
                    onChange={(e) => setMes(e.target.value)}
                    placeholder="Ej: Mayo 2026"
                    className="w-full px-5 py-4 bg-[#F7F5F2] border-none rounded-2xl focus:ring-4 focus:ring-[#5a5a40]/5 focus:bg-white transition-all font-semibold outline-none text-emerald-600"
                  />
                </div>
              </div>

              <button 
                onClick={() => setShowSettingsModal(false)}
                className="w-full py-5 bg-[#1A1A18] text-white rounded-[1.5rem] font-bold text-sm shadow-xl shadow-black/10 hover:bg-black transition-all active:scale-[0.98]"
              >
                Guardar cambios
              </button>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Bulk Confirm Modal */}
      <AnimatePresence>
        {showBulkConfirmModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowBulkConfirmModal(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="relative bg-white w-full max-w-md rounded-[2.5rem] shadow-2xl p-8 md:p-10 flex flex-col gap-8"
            >
              <div className="text-center space-y-3">
                <div className="w-16 h-16 bg-emerald-50 text-emerald-600 rounded-full flex items-center justify-center mx-auto mb-2">
                  <MessageSquare size={32} />
                </div>
                <h3 className="text-2xl font-bold tracking-tight text-[#1A1A18]">Asistente de Envío</h3>
                <p className="text-sm text-[#86867A] font-medium leading-relaxed px-4">
                  Se procesarán {data.length} registros en {new Set(data.map(d => d.telefono)).size} conversaciones únicas por WhatsApp.
                </p>
              </div>
              
              <div className="flex flex-col gap-3">
                <button 
                  onClick={handleBulkSend}
                  className="w-full py-5 bg-emerald-600 text-white rounded-[1.5rem] font-bold text-sm hover:bg-emerald-700 transition-all shadow-xl shadow-emerald-600/10 flex items-center justify-center gap-2 group"
                >
                  <MessageSquare size={18} className="group-hover:scale-110 transition-transform" />
                  Iniciar Envío Automático
                </button>
                <button 
                  onClick={handleBulkCopy}
                  className="w-full py-5 bg-[#1A1A18] text-white rounded-[1.5rem] font-bold text-sm hover:bg-black transition-all shadow-xl shadow-black/10 flex items-center justify-center gap-2"
                >
                  <Copy size={18} />
                  Copiar Todo al Portapapeles
                </button>
                <button 
                  onClick={() => setShowBulkConfirmModal(false)}
                  className="w-full py-4 text-[#86867A] font-bold text-xs uppercase tracking-widest hover:text-[#1A1A18] transition-colors"
                >
                  Regresar al Dashboard
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
      {/* Clear All Confirm Modal */}
      <AnimatePresence>
        {showClearConfirmModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowClearConfirmModal(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="relative bg-white w-full max-w-md rounded-[2.5rem] shadow-2xl p-8 md:p-10 flex flex-col gap-8"
            >
              <div className="flex flex-col items-center text-center space-y-4">
                <div className="w-16 h-16 bg-rose-50 rounded-2xl flex items-center justify-center text-rose-600">
                  <Trash2 size={32} />
                </div>
                <div className="space-y-2">
                  <h3 className="text-xl font-bold tracking-tight text-[#1A1A18]">¿Limpiar Dashboard?</h3>
                  <p className="text-sm text-[#86867A] leading-relaxed">Se eliminarán permanentemente los {data.length} registros cargados. Esta acción no se puede revertir.</p>
                </div>
              </div>
              
              <div className="flex gap-4">
                <button 
                  onClick={() => setShowClearConfirmModal(false)}
                  className="flex-1 py-4 border border-[#E5E0DA] rounded-[1.5rem] font-bold text-sm hover:bg-[#F7F5F2] transition-colors"
                >
                  Cancelar
                </button>
                <button 
                  onClick={handleClearAll}
                  className="flex-1 py-4 bg-rose-600 text-white rounded-[1.5rem] font-bold text-sm hover:bg-rose-700 transition-all shadow-xl shadow-rose-600/10 active:scale-[0.98]"
                >
                  Confirmar
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
      {/* Paste Modal */}
      <AnimatePresence>
        {showPasteModal && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowPasteModal(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ scale: 0.95, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.95, opacity: 0 }}
              className="relative bg-white w-full max-w-3xl rounded-[2.5rem] shadow-2xl p-8 md:p-10 flex flex-col gap-8"
            >
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-2xl font-bold tracking-tight text-[#1A1A18]">Importación rápida</h3>
                  <p className="text-xs text-[#86867A] font-bold uppercase tracking-widest mt-1">Excel or Google Sheets</p>
                </div>
                <button 
                  onClick={() => setShowPasteModal(false)}
                  className="p-2 hover:bg-[#F7F5F2] rounded-full transition-colors text-[#86867A]"
                >
                  <X size={20} />
                </button>
              </div>
              
              <div className="space-y-4">
                <p className="text-xs text-[#86867A] font-medium leading-relaxed">
                  Pega el contenido de tu tabla aquí. El sistema detectará automáticamente los encabezados basados en este orden:
                </p>
                <div className="flex flex-wrap gap-2">
                  {FIXED_HEADERS.slice(0, 8).map((h, i) => (
                    <span key={i} className="px-2 py-1 bg-[#F7F5F2] rounded-lg text-[10px] font-bold text-[#5a5a40] uppercase tracking-tighter">
                      {h}
                    </span>
                  ))}
                  <span className="px-2 py-1 bg-[#F7F5F2] rounded-lg text-[10px] font-bold text-[#5a5a40] uppercase tracking-tighter">...</span>
                </div>
                
                <textarea 
                  value={pasteText}
                  onChange={(e) => setPasteText(e.target.value)}
                  placeholder="Control + V para pegar los datos aquí..."
                  className="w-full h-64 p-6 text-xs font-mono bg-[#FAF9F7] border border-[#E5E0DA] rounded-[2rem] focus:ring-4 focus:ring-[#5a5a40]/5 focus:border-[#5a5a40] outline-none transition-all resize-none shadow-inner"
                />
              </div>

              <div className="flex gap-4">
                <button 
                  onClick={() => setShowPasteModal(false)}
                  className="flex-1 py-4 border border-[#E5E0DA] rounded-[1.5rem] font-bold text-sm hover:bg-[#F7F5F2] transition-colors"
                >
                  Cancelar
                </button>
                <button 
                  onClick={handlePaste}
                  disabled={!pasteText.trim()}
                  className="flex-1 py-4 bg-[#1A1A18] text-white rounded-[1.5rem] font-bold text-sm hover:bg-black disabled:opacity-30 disabled:cursor-not-allowed transition-all shadow-xl shadow-black/10 active:scale-[0.98]"
                >
                  Procesar datos
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}
