export interface Antecedentes {
  antecedentes_familiares: string | null;
  antecedentes_patologicos: string | null;
  antecedentes_personales: string | null;
  antecedentes_quirurgicos: string | null;
  antecedentes_traumaticos: string | null;
  antecedentes_farmacologicos: string | null;
}

export interface MetricasRelevantes {
  presion_arterial: string | null;
  nivel_azucar: string | null;
  hba1c: string | null;
}

export interface HistoriaAnterior {
  id_archivo: number;
  codigo_cita: string;
  nombre_especialidad: string;
  nombre_aliado: string;
  fecha_atencion: string;
  url_archivo: string;
}

export interface AtencionProgramada {
  codigo_cita: string;
  fecha_cita: string;
  nombre_especialidad: string;
  nombre_profesional: string;
  estado_cita: string;
}

export interface ProgEspecialidadActa {
  id: number;
  id_especialidad: number;
  nombre_especialidad: string;
  fecha_inicio: string;
  dias: number;
}

export interface DatosActa {
  id_acta: number;
  id_tipo_identificacion: number;
  identificacion: string;
  tipo_identificacion: string;
  antecedentes: Antecedentes;
  nombre_completo: string;
  fecha_nacimiento: string;
  edad: string;
  genero: string;
  nivel_riesgo: number;
  fecha_proxima_revision: string | null;
  analisis: string | null;
  observaciones_clinicas: string | null;
  observaciones_operativas: string | null;
  id_usuario_acta: number;
  id_estado_actual: number;
  descripcion_riesgo: string | null;
  metricas_relevantes: MetricasRelevantes;
  actas_anteriores: any;
  archivos_paciente: any;
  historias_anteriores: HistoriaAnterior[] | null;
  atenciones_programadas: AtencionProgramada[] | null;
  prog_especialidad_acta: ProgEspecialidadActa[] | null;
  ordenes_servicio: any;
  diagnosticos: any;
  medicamentos_prescritos: any;
}

export interface Patient {
  identificacion: string;
  nombre_completo: string;
  nombre_convenio: string;
  datos_acta: DatosActa | null;
}
