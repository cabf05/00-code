// src/components/ExcelImporter.js - VERSÃO COMPLETA ATUALIZADA (i18next)
import React, { useState, useEffect } from 'react';
import { supabase } from '../utils/supabaseClient';
import { toast } from 'react-hot-toast';
import { 
  FiX, 
  FiDownload, 
  FiUpload, 
  FiCheck, 
  FiAlertCircle, 
  FiFile,
  FiCpu,
  FiChevronDown,
  FiChevronUp
} from 'react-icons/fi';
import ExcelJS from 'exceljs';
import { useTranslation } from 'next-i18next';
import { serverSideTranslations } from 'next-i18next/serverSideTranslations';

const ExcelImporter = ({ 
  onClose, 
  onSuccess, 
  type = 'tarefas', // 'tarefas' ou 'rotinas'
  listas,
  projetos,
  usuarios,
  usuariosListas
}) => {
  const { t } = useTranslation('common');

  const [step, setStep] = useState(1); // 1: Download, 2: Upload, 3: Confirmação
  const [file, setFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [dadosParaImportar, setDadosParaImportar] = useState([]);
  const [errosValidacao, setErrosValidacao] = useState([]);
  const [showAdvancedOptions, setShowAdvancedOptions] = useState(false);

  // Labels / textos reutilizáveis
  const typeLabel = type === 'tarefas' ? t('type.tasks') : t('type.routines');
  const sheetName = type === 'tarefas' ? t('sheet.tasks') : t('sheet.routines');

  // ✅ TIPOS DE RECORRÊNCIA COMPLETOS (igual ao tarefas-rotinas.js)
  const recurrenceTypes = [
    { value: 'daily', label: t('recurrence.daily') },
    { value: 'weekly', label: t('recurrence.weekly') },
    { value: 'monthly', label: t('recurrence.monthly') },
    { value: 'yearly', label: t('recurrence.yearly') },
    { value: 'biweekly', label: t('recurrence.biweekly') },
    { value: 'triweekly', label: t('recurrence.triweekly') },
    { value: 'quadweekly', label: t('recurrence.quadweekly') },
    { value: 'monthly_weekday', label: t('recurrence.monthly_weekday') }
  ];

  // ✅ TIPOS BÁSICOS E AVANÇADOS SEPARADOS
  const tiposRecorrenciaBasicos = [
    { value: 'daily', label: t('recurrence.daily') },
    { value: 'weekly', label: t('recurrence.weekly') },
    { value: 'monthly', label: t('recurrence.monthly') },
    { value: 'yearly', label: t('recurrence.yearly') }
  ];

  const tiposRecorrenciaAvancados = [
    { value: 'biweekly', label: t('recurrence.biweekly') },
    { value: 'triweekly', label: t('recurrence.triweekly') },
    { value: 'quadweekly', label: t('recurrence.quadweekly') },
    { value: 'monthly_weekday', label: t('recurrence.monthly_weekday_short') }
  ];

  // ✅ ORDINAIS PARA PADRÕES MENSIAIS
  const ordinaisMensais = [
    { value: 1, label: t('ordinal.first') },
    { value: 2, label: t('ordinal.second') },
    { value: 3, label: t('ordinal.third') },
    { value: 4, label: t('ordinal.fourth') },
    { value: -1, label: t('ordinal.last') }
  ];

  // ✅ DIAS DA SEMANA
  const diasDaSemana = [
    { valor: 1, nome: t('day.monday'), abrev: t('day.short.monday') },
    { valor: 2, nome: t('day.tuesday'), abrev: t('day.short.tuesday') },
    { valor: 3, nome: t('day.wednesday'), abrev: t('day.short.wednesday') },
    { valor: 4, nome: t('day.thursday'), abrev: t('day.short.thursday') },
    { valor: 5, nome: t('day.friday'), abrev: t('day.short.friday') },
    { valor: 6, nome: t('day.saturday'), abrev: t('day.short.saturday') },
    { valor: 0, nome: t('day.sunday'), abrev: t('day.short.sunday') }
  ];

  // Função para gerar planilha Excel
  const gerarPlanilhaExcel = async () => {
    try {
      setLoading(true);
      
      // Criar workbook
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet(sheetName);
      
      // Definir colunas baseadas no tipo
      if (type === 'tarefas') {
        worksheet.columns = [
          { header: t('headers.task_description'), key: 'content', width: 50 },
          { header: t('headers.list_name_or_id'), key: 'task_list_id', width: 30 },
          { header: t('headers.responsible_name_or_id'), key: 'usuario_id', width: 25 },
          { header: t('headers.deadline'), key: 'date', width: 20 },
          { header: t('headers.status'), key: 'completed', width: 15 },
          { header: t('headers.note'), key: 'note', width: 40 }
        ];
      } else {
        worksheet.columns = [
          { header: t('headers.routine_description'), key: 'content', width: 50 },
          { header: t('headers.list_name_or_id'), key: 'task_list_id', width: 30 },
          { header: t('headers.responsible_name_or_id'), key: 'usuario_id', width: 25 },
          { header: t('headers.recurrence_type'), key: 'recurrence_type', width: 20 },
          { header: t('headers.recurrence_interval'), key: 'recurrence_interval', width: 20 },
          { header: t('headers.recurrence_days'), key: 'recurrence_days', width: 20 },
          { header: t('headers.start_date'), key: 'start_date', width: 20 },
          { header: t('headers.end_date'), key: 'end_date', width: 20 },
          { header: t('headers.persistent'), key: 'persistent', width: 15 },
          { header: t('headers.note'), key: 'note', width: 40 },
          { header: t('headers.weekly_interval'), key: 'weekly_interval', width: 20 },
          { header: t('headers.selected_weekday'), key: 'selected_weekday', width: 20 },
          { header: t('headers.monthly_ordinal'), key: 'monthly_ordinal', width: 20 },
          { header: t('headers.monthly_weekday'), key: 'monthly_weekday', width: 20 }
        ];
      }
      
      // Formatar cabeçalho
      worksheet.getRow(1).font = { bold: true };
      worksheet.getRow(1).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFE6F3FF' }
      };
      
      // Adicionar dados de exemplo
      if (type === 'tarefas') {
        const exemploTarefas = [
          {
            content: t('examples.task1.content'),
            task_list_id: t('examples.task1.list'),
            usuario_id: 'João Silva',
            date: '2024-12-31',
            completed: 'false',
            note: t('examples.task1.note')
          },
          {
            content: t('examples.task2.content'),
            task_list_id: t('examples.task2.list'),
            usuario_id: 'Maria Santos',
            date: '2024-12-15',
            completed: 'true',
            note: ''
          },
          {
            content: t('examples.task3.content'),
            task_list_id: t('examples.task3.list'),
            usuario_id: 'Pedro Costa',
            date: '2024-11-30',
            completed: 'false',
            note: t('examples.task3.note')
          }
        ];
        
        exemploTarefas.forEach(tarefa => {
          worksheet.addRow(tarefa);
        });
      } else {
        const exemploRotinas = [
          {
            content: t('examples.routine1.content'),
            task_list_id: t('examples.routine1.list'),
            usuario_id: 'Pedro Costa',
            recurrence_type: 'daily',
            recurrence_interval: '1',
            recurrence_days: '',
            start_date: '2024-01-01',
            end_date: '',
            persistent: 'true',
            note: t('examples.routine1.note'),
            weekly_interval: '',
            selected_weekday: '',
            monthly_ordinal: '',
            monthly_weekday: ''
          },
          {
            content: t('examples.routine2.content'),
            task_list_id: t('examples.routine2.list'),
            usuario_id: 'Ana Oliveira',
            recurrence_type: 'weekly',
            recurrence_interval: '1',
            recurrence_days: '1,2,3,4,5',
            start_date: '2024-01-01',
            end_date: '2024-12-31',
            persistent: 'true',
            note: t('examples.routine2.note'),
            weekly_interval: '',
            selected_weekday: '',
            monthly_ordinal: '',
            monthly_weekday: ''
          },
          {
            content: t('examples.routine3.content'),
            task_list_id: t('examples.routine3.list'),
            usuario_id: 'Carlos Silva',
            recurrence_type: 'monthly',
            recurrence_interval: '1',
            recurrence_days: '',
            start_date: '2024-01-01',
            end_date: '',
            persistent: 'true',
            note: t('examples.routine3.note'),
            weekly_interval: '',
            selected_weekday: '',
            monthly_ordinal: '',
            monthly_weekday: ''
          },
          {
            content: t('examples.routine4.content'),
            task_list_id: t('examples.routine4.list'),
            usuario_id: 'Ana Oliveira',
            recurrence_type: 'biweekly',
            recurrence_interval: '1',
            recurrence_days: '',
            start_date: '2024-01-01',
            end_date: '',
            persistent: 'true',
            note: t('examples.routine4.note'),
            weekly_interval: '2',
            selected_weekday: '3',
            monthly_ordinal: '',
            monthly_weekday: ''
          },
          {
            content: t('examples.routine5.content'),
            task_list_id: t('examples.routine5.list'),
            usuario_id: 'Ana Oliveira',
            recurrence_type: 'monthly_weekday',
            recurrence_interval: '1',
            recurrence_days: '',
            start_date: '2024-01-01',
            end_date: '',
            persistent: 'true',
            note: t('examples.routine5.note'),
            weekly_interval: '',
            selected_weekday: '',
            monthly_ordinal: '1',
            monthly_weekday: '1'
          }
        ];
        
        exemploRotinas.forEach(rotina => {
          worksheet.addRow(rotina);
        });
      }
      
      // Adicionar bordas
      worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell) => {
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      });
      
      // Adicionar instruções em uma nova aba
      const instrucoes = workbook.addWorksheet(t('instructions.sheet_header', { type: typeLabel.toUpperCase ? typeLabel.toUpperCase() : String(typeLabel).toUpperCase() }));
      
      instrucoes.addRow([t('instructions.sheet_header_plain', { type: typeLabel.toUpperCase ? typeLabel.toUpperCase() : String(typeLabel).toUpperCase() })]);
      instrucoes.addRow([]);
      instrucoes.addRow([t('instructions.required_title')]);
      
      if (type === 'tarefas') {
        instrucoes.addRow([t('instructions.task_description')]);
        instrucoes.addRow([t('instructions.list_field')]);
        instrucoes.addRow([t('instructions.responsible_field')]);
        instrucoes.addRow([]);
        instrucoes.addRow([t('instructions.optional_title')]);
        instrucoes.addRow([t('instructions.deadline_format')]);
        instrucoes.addRow([t('instructions.status_format')]);
        instrucoes.addRow([t('instructions.note_field')]);
      } else {
        instrucoes.addRow([t('instructions.routine_description')]);
        instrucoes.addRow([t('instructions.list_field')]);
        instrucoes.addRow([t('instructions.responsible_field')]);
        instrucoes.addRow([t('instructions.recurrence_type_field')]);
        instrucoes.addRow([t('instructions.start_date_format')]);
        instrucoes.addRow([]);
        instrucoes.addRow([t('instructions.optional_title')]);
        instrucoes.addRow([t('instructions.recurrence_interval')]);
        instrucoes.addRow([t('instructions.recurrence_days')]);
        instrucoes.addRow([t('instructions.end_date_format')]);
        instrucoes.addRow([t('instructions.persistent_field')]);
        instrucoes.addRow([t('instructions.note_field')]);
        instrucoes.addRow([]);
        instrucoes.addRow([t('instructions.advanced_title')]);
        instrucoes.addRow([t('instructions.weekly_interval')]);
        instrucoes.addRow([t('instructions.weekday_field')]);
        instrucoes.addRow([t('instructions.monthly_ordinal')]);
        instrucoes.addRow([t('instructions.monthly_weekday')]);
      }
      
      instrucoes.addRow([]);
      instrucoes.addRow([t('instructions.available_lists')]);
      Object.entries(listas).forEach(([id, lista]) => {
        const projetoName = projetos[lista.projeto_id] || t('general.project_not_found');
        instrucoes.addRow([t('instructions.list_item', { id, name: lista.nome, project: projetoName })]);
      });
      
      instrucoes.addRow([]);
      instrucoes.addRow([t('instructions.available_users')]);
      Object.entries(usuarios).forEach(([id, nome]) => {
        instrucoes.addRow([t('instructions.user_item', { id, name: nome })]);
      });
      
      instrucoes.addRow([]);
      instrucoes.addRow([t('instructions.notes_title')]);
      instrucoes.addRow([t('instructions.note_use_list_or_id')]);
      instrucoes.addRow([t('instructions.note_use_user_or_id')]);
      instrucoes.addRow([t('instructions.note_date_format')]);
      instrucoes.addRow([t('instructions.note_remove_example_rows')]);
      instrucoes.addRow([t('instructions.note_required_fields')]);
      instrucoes.addRow([t('instructions.note_advanced_types')]);
      
      // Formatar instruções
      instrucoes.getRow(1).font = { bold: true, size: 16 };
      instrucoes.getColumn(1).width = 100;
      
      // Formatar seções
      const sections = [
        { row: 1, bgColor: 'FF2F75B6', textColor: 'FFFFFFFF' },
        { row: 3, bgColor: 'FFFCE4D6', textColor: 'FF8B4513' },
        { row: type === 'tarefas' ? 9 : 11, bgColor: 'FFE7F5E6', textColor: 'FF2E8B57' },
        { row: type === 'tarefas' ? 14 : 35, bgColor: 'FFFFF2CC', textColor: 'FF8B6914' }
      ];
      
      sections.forEach(section => {
        if (instrucoes.getRow(section.row)) {
          instrucoes.getRow(section.row).fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: section.bgColor }
          };
          instrucoes.getRow(section.row).font = {
            bold: true,
            color: { argb: section.textColor }
          };
        }
      });
      
      // Gerar buffer e baixar
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      
      // Criar link de download
      const link = document.createElement('a');
      link.href = url;
      const timestamp = new Date().toISOString().split('T')[0];
      const typeForFile = type === 'tarefas' ? t('file.type.tasks') : t('file.type.routines');
      link.download = t('file.name', { type: typeForFile, timestamp });
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      // Limpar URL
      window.URL.revokeObjectURL(url);
      
      toast.success(t('toast.download_success', { type: typeLabel }));
      setStep(2);
      
    } catch (error) {
      console.error(t('console.error_generate_workbook'), error);
      toast.error(t('toast.error_generate_template'));
    } finally {
      setLoading(false);
    }
  };

  // Função para processar arquivo Excel
  const processarArquivoExcel = async () => {
    if (!file) {
      toast.error(t('toast.select_file'));
      return;
    }

    try {
      setLoading(true);
      
      const buffer = await file.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      
      const worksheet = workbook.getWorksheet(type === 'tarefas' ? sheetName : sheetName);
      if (!worksheet) {
        throw new Error(t('errors.sheet_not_found', { sheetName }));
      }
      
      const dadosProcessados = [];
      const erros = [];
      
      // Processar cada linha (começando da linha 2, pulando o cabeçalho)
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Pular cabeçalho
        
        const dadosLinha = {};
        
        if (type === 'tarefas') {
          dadosLinha.content = row.getCell(1).value;
          dadosLinha.task_list_id = row.getCell(2).value;
          dadosLinha.usuario_id = row.getCell(3).value;
          dadosLinha.date = row.getCell(4).value;
          dadosLinha.completed = row.getCell(5).value;
          dadosLinha.note = row.getCell(6).value;
        } else {
          dadosLinha.content = row.getCell(1).value;
          dadosLinha.task_list_id = row.getCell(2).value;
          dadosLinha.usuario_id = row.getCell(3).value;
          datosLinhaDummy: null; // (placeholder if needed)
          dadosLinha.recurrence_type = row.getCell(4).value;
          dadosLinha.recurrence_interval = row.getCell(5).value;
          dadosLinha.recurrence_days = row.getCell(6).value;
          dadosLinha.start_date = row.getCell(7).value;
          dadosLinha.end_date = row.getCell(8).value;
          dadosLinha.persistent = row.getCell(9).value;
          dadosLinha.note = row.getCell(10).value;
          dadosLinha.weekly_interval = row.getCell(11).value;
          dadosLinha.selected_weekday = row.getCell(12).value;
          dadosLinha.monthly_ordinal = row.getCell(13).value;
          dadosLinha.monthly_weekday = row.getCell(14).value;
        }
        
        // Validações básicas
        if (!dadosLinha.content || String(dadosLinha.content).trim() === '') {
          erros.push(t('validation.description_required', { row: rowNumber }));
          return;
        }
        
        if (!dadosLinha.task_list_id) {
          erros.push(t('validation.list_required', { row: rowNumber }));
          return;
        }
        
        if (!dadosLinha.usuario_id) {
          erros.push(t('validation.responsible_required', { row: rowNumber }));
          return;
        }
        
        // Validações específicas para rotinas
        if (type === 'rotinas') {
          if (!dadosLinha.recurrence_type) {
            erros.push(t('validation.recurrence_required', { row: rowNumber }));
            return;
          }
          
          if (!recurrenceTypes.find(rt => rt.value === dadosLinha.recurrence_type)) {
            const allowed = recurrenceTypes.map(rt => rt.value).join(', ');
            erros.push(t('validation.recurrence_invalid', { row: rowNumber, allowed }));
            return;
          }
          
          if (!dadosLinha.start_date) {
            erros.push(t('validation.start_date_required', { row: rowNumber }));
            return;
          }
          
          // ✅ VALIDAÇÕES PARA TIPOS AVANÇADOS
          if (['biweekly', 'triweekly', 'quadweekly'].includes(dadosLinha.recurrence_type)) {
            if (!dadosLinha.selected_weekday) {
              erros.push(t('validation.weekday_required', { row: rowNumber, type: dadosLinha.recurrence_type }));
              return;
            }
            
            if (!dadosLinha.weekly_interval) {
              // Definir intervalo padrão baseado no tipo
              if (dadosLinha.recurrence_type === 'biweekly') dadosLinha.weekly_interval = 2;
              else if (dadosLinha.recurrence_type === 'triweekly') dadosLinha.weekly_interval = 3;
              else if (dadosLinha.recurrence_type === 'quadweekly') dadosLinha.weekly_interval = 4;
            }
          }
          
          if (dadosLinha.recurrence_type === 'monthly_weekday') {
            if (!dadosLinha.monthly_ordinal || !dadosLinha.monthly_weekday) {
              erros.push(t('validation.monthly_required', { row: rowNumber }));
              return;
            }
          }
        }
        
        // Converter e validar task_list_id
        let listaId = null;
        const listaValue = String(dadosLinha.task_list_id).trim();
        
        // Tentar como ID direto
        if (listas[listaValue]) {
          listaId = parseInt(listaValue);
        } else {
          // Buscar por nome
          const foundLista = Object.entries(listas).find(([id, lista]) => 
            lista.nome.toLowerCase() === listaValue.toLowerCase()
          );
          if (foundLista) {
            listaId = parseInt(foundLista[0]);
          } else {
            erros.push(t('validation.list_not_found', { row: rowNumber, lista: listaValue }));
            return;
          }
        }
        
        // Converter e validar usuario_id
        let usuarioId = null;
        const usuarioValue = String(dadosLinha.usuario_id).trim();
        
        // Tentar como ID direto
        if (usuarios[usuarioValue]) {
          usuarioId = usuarioValue;
        } else {
          // Buscar por nome
          const foundUsuario = Object.entries(usuarios).find(([id, nome]) => 
            nome.toLowerCase() === usuarioValue.toLowerCase()
          );
          if (foundUsuario) {
            usuarioId = foundUsuario[0];
          } else {
            erros.push(t('validation.user_not_found', { row: rowNumber, usuario: usuarioValue }));
            return;
          }
        }
        
        // Verificar se usuário tem acesso à lista
        const usuariosDaLista = usuariosListas[listaId] || [];
        if (!usuariosDaLista.includes(usuarioId)) {
          erros.push(t('validation.user_no_access', { row: rowNumber, usuarioName: usuarios[usuarioId], listaName: listas[listaId].nome }));
          return;
        }
        
        // Processar dados específicos por tipo
        const dadosFinais = {
          content: String(dadosLinha.content).trim(),
          task_list_id: listaId,
          usuario_id: usuarioId
        };
        
        if (type === 'tarefas') {
          // Processar data
          if (dadosLinha.date) {
            if (dadosLinha.date instanceof Date) {
              dadosFinais.date = dadosLinha.date.toISOString().split('T')[0];
            } else if (typeof dadosLinha.date === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dadosLinha.date)) {
              dadosFinais.date = dadosLinha.date;
            } else {
              erros.push(t('validation.date_format', { row: rowNumber }));
              return;
            }
          }
          
          // Processar completed
          dadosFinais.completed = false;
          if (dadosLinha.completed) {
            const completedStr = String(dadosLinha.completed).toLowerCase().trim();
            dadosFinais.completed = ['true', '1', 'sim', 'concluída', 'yes'].includes(completedStr);
          }
          
          // Processar nota
          if (dadosLinha.note) {
            dadosFinais.note = String(dadosLinha.note).trim();
          }
        } else {
          // Processar dados de rotina
          dadosFinais.recurrence_type = dadosLinha.recurrence_type;
          dadosFinais.recurrence_interval = parseInt(dadosLinha.recurrence_interval) || 1;
          
          // Processar dias da semana
          if (dadosLinha.recurrence_days && dadosLinha.recurrence_type === 'weekly') {
            const days = String(dadosLinha.recurrence_days)
              .split(',')
              .map(d => parseInt(d.trim()))
              .filter(d => !isNaN(d) && d >= 0 && d <= 6);
            dadosFinais.recurrence_days = days.length > 0 ? days : null;
          } else {
            dadosFinais.recurrence_days = null;
          }
          
          // Processar datas
          if (dadosLinha.start_date instanceof Date) {
            dadosFinais.start_date = dadosLinha.start_date.toISOString().split('T')[0];
          } else if (typeof dadosLinha.start_date === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dadosLinha.start_date)) {
            dadosFinais.start_date = dadosLinha.start_date;
          } else {
            erros.push(t('validation.start_date_format', { row: rowNumber }));
            return;
          }
          
          if (dadosLinha.end_date) {
            if (dadosLinha.end_date instanceof Date) {
              dadosFinais.end_date = dadosLinha.end_date.toISOString().split('T')[0];
            } else if (typeof dadosLinha.end_date === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dadosLinha.end_date)) {
              dadosFinais.end_date = dadosLinha.end_date;
            } else {
              erros.push(t('validation.end_date_format', { row: rowNumber }));
              return;
            }
          }
          
          // Processar persistente
          dadosFinais.persistent = true;
          if (dadosLinha.persistent) {
            const persistentStr = String(dadosLinha.persistent).toLowerCase().trim();
            dadosFinais.persistent = !['false', '0', 'não', 'nao', 'no'].includes(persistentStr);
          }
          
          // Processar nota
          if (dadosLinha.note) {
            dadosFinais.note = String(dadosLinha.note).trim();
          }
          
          // ✅ PROCESSAR CAMPOS AVANÇADOS
          if (dadosLinha.weekly_interval) {
            dadosFinais.weekly_interval = parseInt(dadosLinha.weekly_interval) || 
              (dadosLinha.recurrence_type === 'biweekly' ? 2 : 
               dadosLinha.recurrence_type === 'triweekly' ? 3 : 4);
          }
          
          if (dadosLinha.selected_weekday) {
            dadosFinais.selected_weekday = parseInt(dadosLinha.selected_weekday);
          }
          
          if (dadosLinha.monthly_ordinal) {
            dadosFinais.monthly_ordinal = parseInt(dadosLinha.monthly_ordinal);
          }
          
          if (dadosLinha.monthly_weekday) {
            dadosFinais.monthly_weekday = parseInt(dadosLinha.monthly_weekday);
          }
        }
        
        dadosProcessados.push({
          ...dadosFinais,
          _rowIndex: rowNumber
        });
      });
      
      if (erros.length > 0) {
        setErrosValidacao(erros);
        toast.error(t('toast.sheet_errors', { count: erros.length }));
        setLoading(false);
        return;
      }
      
      setDadosParaImportar(dadosProcessados);
      setStep(3);
      toast.success(t('toast.process_success', { count: dadosProcessados.length }));
      
    } catch (error) {
      console.error(t('console.error_process_file'), error);
      toast.error(t('toast.error_process_file'));
    } finally {
      setLoading(false);
    }
  };

  // Função para executar importação
  const executarImportacao = async () => {
    try {
      setLoading(true);
      
      let sucessos = 0;
      let falhas = 0;
      
      const tabela = type === 'tarefas' ? 'tasks' : 'routine_tasks';
      
      for (const item of dadosParaImportar) {
        try {
          const { _rowIndex, ...dadosInsert } = item;
          
          const { error } = await supabase
            .from(tabela)
            .insert(dadosInsert);
            
          if (error) throw error;
          sucessos++;
        } catch (error) {
          console.error(t('console.error_insert_row', { row: item._rowIndex }), error);
          falhas++;
        }
      }
      
      if (sucessos > 0) {
        toast.success(t('toast.import_success', { count: sucessos, type: typeLabel }));
      }
      
      if (falhas > 0) {
        toast.error(t('toast.import_failures', { count: falhas, type: typeLabel }));
      }
      
      if (onSuccess) {
        onSuccess();
      }
      
      onClose();
      
    } catch (error) {
      console.error(t('console.error_import'), error);
      toast.error(t('toast.error_import'));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50">
      <div className="bg-white p-6 rounded-lg max-w-5xl w-full max-h-[90vh] overflow-y-auto">
        <div className="flex justify-between items-center mb-4">
          <h2 className="text-xl font-bold">
            {t('modal.title', { type: typeLabel })}
          </h2>
          <button 
            onClick={onClose}
            className="text-gray-400 hover:text-gray-600"
            aria-label={t('buttons.close')}
          >
            <FiX className="h-6 w-6" />
          </button>
        </div>

        {/* Indicador de progresso */}
        <div className="flex items-center justify-center mb-6">
          <div className="flex items-center space-x-4">
            <div className={`flex items-center ${step >= 1 ? 'text-blue-600' : 'text-gray-400'}`}>
              <div className={`w-8 h-8 rounded-full border-2 flex items-center justify-center ${
                step >= 1 ? 'border-blue-600 bg-blue-50' : 'border-gray-300'
              }`}>
                1
              </div>
              <span className="ml-2 font-medium">{t('steps.download')}</span>
            </div>
            
            <div className={`w-8 h-0.5 ${step >= 2 ? 'bg-blue-600' : 'bg-gray-300'}`}></div>
            
            <div className={`flex items-center ${step >= 2 ? 'text-blue-600' : 'text-gray-400'}`}>
              <div className={`w-8 h-8 rounded-full border-2 flex items-center justify-center ${
                step >= 2 ? 'border-blue-600 bg-blue-50' : 'border-gray-300'
              }`}>
                2
              </div>
              <span className="ml-2 font-medium">{t('steps.upload')}</span>
            </div>
            
            <div className={`w-8 h-0.5 ${step >= 3 ? 'bg-blue-600' : 'bg-gray-300'}`}></div>
            
            <div className={`flex items-center ${step >= 3 ? 'text-blue-600' : 'text-gray-400'}`}>
              <div className={`w-8 h-8 rounded-full border-2 flex items-center justify-center ${
                step >= 3 ? 'border-blue-600 bg-blue-50' : 'border-gray-300'
              }`}>
                3
              </div>
              <span className="ml-2 font-medium">{t('steps.confirmation')}</span>
            </div>
          </div>
        </div>

        {/* Conteúdo por etapa */}
        {step === 1 && (
          <div className="space-y-4">
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
              <h3 className="font-semibold text-blue-800 mb-2">{t('instructions.how_it_works_title')}</h3>
              <ol className="list-decimal list-inside space-y-1 text-blue-700">
                <li>{t('instructions.step1')}</li>
                <li>{t('instructions.step2')}</li>
                <li>{t('instructions.step3')}</li>
                <li>{t('instructions.step4')}</li>
              </ol>
            </div>

            <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
              <h3 className="font-semibold text-yellow-800 mb-2">{t('important.title')}</h3>
              <ul className="list-disc list-inside space-y-1 text-yellow-700">
                <li>{t('important.keep_columns')}</li>
                <li>{t('important.required_marked')}</li>
                <li>{t('important.advanced_fill')}</li>
                <li>{t('important.check_user_access')}</li>
              </ul>
            </div>

            <button
              onClick={gerarPlanilhaExcel}
              disabled={loading}
              className="w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-3 px-4 rounded-lg flex items-center justify-center disabled:opacity-50"
            >
              {loading ? (
                <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-white"></div>
              ) : (
                <>
                  <FiDownload className="mr-2" />
                  {t('buttons.download_template')}
                </>
              )}
            </button>
          </div>
        )}

        {step === 2 && (
          <div className="space-y-4">
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center">
              <input
                type="file"
                id="excel-file"
                accept=".xlsx,.xls"
                onChange={(e) => setFile(e.target.files[0])}
                className="hidden"
              />
              
              <label htmlFor="excel-file" className="cursor-pointer">
                {file ? (
                  <div className="text-green-600">
                    <FiCheck className="h-12 w-12 mx-auto mb-2" />
                    <p className="font-medium">{file.name}</p>
                    <p className="text-sm text-gray-500">{t('file.click_to_select_another')}</p>
                  </div>
                ) : (
                  <div className="text-gray-400">
                    <FiUpload className="h-12 w-12 mx-auto mb-2" />
                    <p className="font-medium">{t('file.click_to_select')}</p>
                    <p className="text-sm">{t('file.supported_formats')}</p>
                  </div>
                )}
              </label>
            </div>

            {errosValidacao.length > 0 && (
              <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                <h4 className="font-semibold text-red-800 mb-2">{t('errors.title')}</h4>
                <div className="max-h-32 overflow-y-auto">
                  {errosValidacao.map((erro, index) => (
                    <p key={index} className="text-red-700 text-sm mb-1">
                      {erro}
                    </p>
                  ))}
                </div>
              </div>
            )}

            <div className="flex space-x-3">
              <button
                onClick={() => setStep(1)}
                className="flex-1 bg-gray-300 hover:bg-gray-400 text-gray-800 font-medium py-2 px-4 rounded-lg"
              >
                {t('buttons.back')}
              </button>
              <button
                onClick={processarArquivoExcel}
                disabled={!file || loading}
                className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-medium py-2 px-4 rounded-lg disabled:opacity-50"
              >
                {loading ? t('buttons.processing') : t('buttons.process_file')}
              </button>
            </div>
          </div>
        )}

        {step === 3 && (
          <div className="space-y-4">
            <div className="bg-green-50 border border-green-200 rounded-lg p-4">
              <div className="flex items-center">
                <FiCheck className="h-5 w-5 text-green-600 mr-2" />
                <span className="text-green-800 font-medium">
                  {t('confirmation.ready_to_import', { count: dadosParaImportar.length, type: typeLabel })}
                </span>
              </div>
            </div>

            {/* Visualização dos dados */}
            <div className="border rounded-lg overflow-hidden">
              <div className="bg-gray-50 px-4 py-2 font-medium">
                {t('table.data_title', { count: dadosParaImportar.length })}
              </div>
              <div className="max-h-64 overflow-y-auto">
                <table className="w-full text-sm">
                  <thead className="bg-gray-100">
                    <tr>
                      <th className="px-3 py-2 text-left">{t('table.headers.description')}</th>
                      <th className="px-3 py-2 text-left">{t('table.headers.list')}</th>
                      <th className="px-3 py-2 text-left">{t('table.headers.responsible')}</th>
                      {type === 'rotinas' && (
                        <th className="px-3 py-2 text-left">{t('table.headers.recurrence')}</th>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {dadosParaImportar.slice(0, 10).map((item, index) => (
                      <tr key={index} className="border-t">
                        <td className="px-3 py-2">{item.content}</td>
                        <td className="px-3 py-2">{listas[item.task_list_id]?.nome}</td>
                        <td className="px-3 py-2">{usuarios[item.usuario_id]}</td>
                        {type === 'rotinas' && (
                          <td className="px-3 py-2">
                            {recurrenceTypes.find(rt => rt.value === item.recurrence_type)?.label}
                            {item.recurrence_interval > 1 && ` (${t('table.every_interval', { interval: item.recurrence_interval })})`}
                          </td>
                        )}
                      </tr>
                    ))}
                    {dadosParaImportar.length > 10 && (
                      <tr>
                        <td colSpan={type === 'tarefas' ? 3 : 4} className="px-3 py-2 text-center text-gray-500">
                          {t('table.more', { count: dadosParaImportar.length - 10 })}
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="flex space-x-3">
              <button
                onClick={() => setStep(2)}
                className="flex-1 bg-gray-300 hover:bg-gray-400 text-gray-800 font-medium py-2 px-4 rounded-lg"
              >
                {t('buttons.back')}
              </button>
              <button
                onClick={executarImportacao}
                disabled={loading}
                className="flex-1 bg-green-600 hover:bg-green-700 text-white font-medium py-2 px-4 rounded-lg disabled:opacity-50"
              >
                {loading ? t('buttons.importing') : t('buttons.confirm_import')}
              </button>
            </div>
          </div>
        )}

        {/* Seção de opções avançadas */}
        {type === 'rotinas' && step === 1 && (
          <div className="mt-6 border-t pt-4">
            <button
              onClick={() => setShowAdvancedOptions(!showAdvancedOptions)}
              className="flex items-center text-blue-600 hover:text-blue-800 font-medium"
            >
              <FiCpu className="mr-2" />
              {t('advanced.button')}
              {showAdvancedOptions ? (
                <FiChevronUp className="ml-2" />
              ) : (
                <FiChevronDown className="ml-2" />
              )}
            </button>

            {showAdvancedOptions && (
              <div className="mt-3 bg-gray-50 p-4 rounded-lg">
                <h4 className="font-semibold mb-3">{t('advanced.supported_title')}</h4>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div>
                    <h5 className="font-medium text-blue-800 mb-2">{t('advanced.basic_title')}</h5>
                    <ul className="space-y-1 text-sm">
                      {tiposRecorrenciaBasicos.map(tipo => (
                        <li key={tipo.value} className="flex items-center">
                          <FiCheck className="h-3 w-3 text-green-600 mr-2" />
                          <span className="font-medium">{tipo.label}:</span>
                          <span className="ml-1 text-gray-600">{tipo.value}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                  
                  <div>
                    <h5 className="font-medium text-purple-800 mb-2">{t('advanced.advanced_title')}</h5>
                    <ul className="space-y-1 text-sm">
                      {tiposRecorrenciaAvancados.map(tipo => (
                        <li key={tipo.value} className="flex items-center">
                          <FiCpu className="h-3 w-3 text-purple-600 mr-2" />
                          <span className="font-medium">{tipo.label}:</span>
                          <span className="ml-1 text-gray-600">{tipo.value}</span>
                        </li>
                      ))}
                    </ul>
                  </div>
                </div>

                <div className="mt-4 p-3 bg-white rounded border">
                  <h6 className="font-medium text-orange-800 mb-2">{t('examples.title')}</h6>
                  <div className="text-xs space-y-2">
                    <p><strong>biweekly:</strong> {t('examples.biweekly')}</p>
                    <p><strong>monthly_weekday:</strong> {t('examples.monthly_weekday')}</p>
                    <p><strong>quadweekly:</strong> {t('examples.quadweekly')}</p>
                  </div>
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default ExcelImporter;
