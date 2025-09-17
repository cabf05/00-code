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
      if (!worksh
