require('dotenv').config({ path: 'config' });

var SOURCE_FILE = process.env.SOURCE_FILE;
var OUTPUT_FILE = process.env.OUTPUT_FILE;
var WORKSHEET = process.env.WORKSHEET;
var CREATED_BY_COLUMN = process.env.CREATED_BY_COLUMN;
var TITLE_COLUMN = process.env.TITLE_COLUMN;
var PRIMARY_LABELS_COLUMN = process.env.PRIMARY_LABELS_COLUMN;
var CLOSED_AT = process.env.CLOSED_AT;
var LABELS_COLUMN = process.env.LABELS_COLUMN;
var CLIENTS_COLUMN = process.env.CLIENTS_COLUMN;
var CREATED_AT = process.env.CREATED_AT;
var CARD_ASSIGNEES = process.env.CARD_ASSIGNEES;
var CATEGORIA = process.env.CATEGORIA;
var SERVICE_LINE = process.env.SERVICE_LINE;
var PROBLEMA_REPORTADO = process.env.PROBLEMA_REPORTADO;
var ANALISE_ACIONAMENTO = process.env.ANALISE_ACIONAMENTO;
var LABELS_ALEATORIAS = process.env.LABELS_ALEATORIAS;
var ACAO_ISM = process.env.ACAO_ISM;
var CANAL_ACIONAMENTO = process.env.CANAL_ACIONAMENTO;
var SOLICITACOES = process.env.SOLICITACOES;
var QUEM_VOCE_ACIONOU = process.env.QUEM_VOCE_ACIONOU;
var QUEM_TE_ACIONOU = process.env.QUEM_TE_ACIONOU;
var LABELS_CHAMADOS_INDEVIDOS = process.env.LABELS_CHAMADOS_INDEVIDOS;
var STORE_CREATED_BY_COLUMN = process.env.STORE_CREATED_BY_COLUMN;
var STORE_SHIFT = process.env.STORE_SHIFT;
var STORE_CLIENT_COLUMN = process.env.STORE_CLIENT_COLUMN;
var STORE_TITLE_COLUMN = process.env.STORE_TITLE_COLUMN;
var STORE_PRIMARY_LABELS_COLUMN = process.env.STORE_PRIMARY_LABELS_COLUMN;
var STORE_TYPE_COLUMN = process.env.STORE_TYPE_COLUMN;
var STORE_SEVERITY_COLUNM = process.env.STORE_SEVERITY_COLUNM;
var STORE_CLOSED_AT = process.env.STORE_CLOSED_AT;
var STORE_WEEK_DAY = process.env.STORE_WEEK_DAY;
var STORE_MONTH = process.env.STORE_MONTH;
var STORE_DAY = process.env.STORE_DAY;
var STORE_YEAR = process.env.STORE_YEAR;
var QUANTIDADE_TICKETS_PER_USER = process.env.QUANTIDADE_TICKETS_PER_USER;
var SEV_SUMMARY_LABELS = process.env.SEV_SUMMARY_LABELS;
var SEV_SUMMARY_VALUES = process.env.SEV_SUMMARY_VALUES;
var SEV_SUMMARY_CLIENT_NAME = process.env.SEV_SUMMARY_CLIENT_NAME;
var SEV_SUMMARY_CLIENT_SEV1 = process.env.SEV_SUMMARY_CLIENT_SEV1;
var SEV_SUMMARY_CLIENT_SEV2 = process.env.SEV_SUMMARY_CLIENT_SEV2;
var SEV_SUMMARY_CLIENT_SEV3 = process.env.SEV_SUMMARY_CLIENT_SEV3;
var SEV_SUMMARY_CLIENT_SEV4 = process.env.SEV_SUMMARY_CLIENT_SEV4;
var STORE_WORKED_HOURS = process.env.STORE_WORKED_HOURS;
var STORE_QUANTIDADE_TICKETS = process.env.STORE_QUANTIDADE_TICKETS;
var OPERATIONAL_LEAD_TIME = process.env.OPERATIONAL_LEAD_TIME;
var TOTAL_WAITING_TIME = process.env.TOTAL_WAITING_TIME;
var HORARIO_PICO = process.env.HORARIO_PICO;
var TRIBE = process.env.TRIBE;
var FLOW_EFFICIENCY = process.env.FLOW_EFFICIENCY;
var HORARIO_INCIDENTE = process.env.HORARIO_INCIDENTE;
var SLA_TICKET = process.env.SLA_TICKET;
var HORARIO_ACIONAMENTO = process.env.HORARIO_ACIONAMENTO;
var ISM_SOLICITOU = process.env.ISM_SOLICITOU;
var SLA_TICKET_VENCIDO = process.env.SLA_TICKET_VENCIDO
var TEMPO_ATENDIMENTO = process.env.TEMPO_ATENDIMENTO
var ANALISE_PRAZO_ACIONAMENTO = process.env.ANALISE_PRAZO_ACIONAMENTO

module.exports = {
    SOURCE_FILE,
    OUTPUT_FILE,
    WORKSHEET,
    CREATED_BY_COLUMN,
    TITLE_COLUMN,
    PRIMARY_LABELS_COLUMN,
    CLOSED_AT,
    LABELS_COLUMN,
    CLIENTS_COLUMN,
    CREATED_AT,
    CARD_ASSIGNEES,
    CATEGORIA,
    SERVICE_LINE,
    PROBLEMA_REPORTADO,
    ANALISE_ACIONAMENTO,
    LABELS_ALEATORIAS,
    ACAO_ISM,
    CANAL_ACIONAMENTO,
    SOLICITACOES,
    QUEM_VOCE_ACIONOU,
    QUEM_TE_ACIONOU,
    LABELS_CHAMADOS_INDEVIDOS,
    STORE_CREATED_BY_COLUMN,
    STORE_SHIFT,
    STORE_CLIENT_COLUMN,
    STORE_TITLE_COLUMN,
    STORE_PRIMARY_LABELS_COLUMN,
    STORE_TYPE_COLUMN,
    STORE_SEVERITY_COLUNM,
    STORE_CLOSED_AT,
    STORE_WEEK_DAY,
    STORE_MONTH,
    STORE_DAY,
    STORE_YEAR,
    QUANTIDADE_TICKETS_PER_USER,
    SEV_SUMMARY_LABELS,
    SEV_SUMMARY_VALUES,
    SEV_SUMMARY_CLIENT_NAME,
    SEV_SUMMARY_CLIENT_SEV1,
    SEV_SUMMARY_CLIENT_SEV2,
    SEV_SUMMARY_CLIENT_SEV3,
    SEV_SUMMARY_CLIENT_SEV4,
    STORE_WORKED_HOURS,
    STORE_QUANTIDADE_TICKETS,
    OPERATIONAL_LEAD_TIME,
    TOTAL_WAITING_TIME,
    HORARIO_PICO,
    TRIBE,
    FLOW_EFFICIENCY,
    HORARIO_INCIDENTE,
    SLA_TICKET,
    HORARIO_ACIONAMENTO,
    ISM_SOLICITOU,
    SLA_TICKET_VENCIDO,
    TEMPO_ATENDIMENTO,
    ANALISE_PRAZO_ACIONAMENTO
}