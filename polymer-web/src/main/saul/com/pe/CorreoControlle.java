/**
 * Resumen.          :  asjdjsdfjasn sldjnasflasfn                         
 * Objeto            :  
 * Descripción       :  
 * Fecha de Creación :  
 * PE de Creación    :  
 * Autor             :  
 * ---------------------------------------------------------------------------------------------------------------
* Modificaciones
* Motivo                      Fecha             Nombre                         Descripción
* ---------------------------------------------------------------------------------------------------------------
*/
package pe.com.gmd.sisrec.web.controller;

import static java.text.MessageFormat.format;

import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;

import javax.faces.context.FacesContext;
import javax.faces.model.SelectItem;
import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletResponse;

import org.apache.chemistry.opencmis.client.api.Session;
import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.Region;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Controller;

import pe.com.gmd.seguridad.services.UsuarioService;
import pe.com.gmd.sisrec.core.entity.DetalleGenerico;
import pe.com.gmd.sisrec.core.entity.Empleado;
import pe.com.gmd.sisrec.core.entity.EquipoTrabajo;
import pe.com.gmd.sisrec.core.entity.FrenteAtencion;
import pe.com.gmd.sisrec.core.entity.ParametroReporte;
import pe.com.gmd.sisrec.core.entity.Reclamo;
import pe.com.gmd.sisrec.core.entity.UsuarioReporte;
import pe.com.gmd.sisrec.core.service.AlfrescoService;
import pe.com.gmd.sisrec.core.service.FrenteAtencionService;
import pe.com.gmd.sisrec.core.service.GenericoService;
import pe.com.gmd.sisrec.core.service.ParametroReporteService;
import pe.com.gmd.sisrec.core.service.ReporteService;
import pe.com.gmd.sisrec.core.service.UsuarioReporteService;
import pe.com.gmd.sisrec.core.util.Constantes;
import pe.com.gmd.sisrec.core.util.CoreUtil;
import pe.com.gmd.sisrec.core.util.TablasQueryEnumSisRec;
import pe.com.gmd.sisrec.web.util.FechaUtil;
import pe.com.gmd.sisrec.web.util.Funciones;
import pe.com.gmd.sisrec.web.util.UDocumentoExcel;
import pe.com.gmd.sisrec.web.util.WebUtil;
import pe.com.gmd.util.exception.GmdException;
import pe.com.gmd.util.exception.MensajeExceptionUtil;
import pe.com.gmd.util.properties.PropiedadesUtil;

@Scope("session")
@Controller(value="reportesController")


public class ReportesController implements Serializable{
		
	
	
	private static final long serialVersionUID = 1L;
	private static final Logger LOGGER = Logger.getLogger(ReportesController.class);	
	
	
	@Autowired
	private AlfrescoService alfrescoService;
	@Autowired
	private UsuarioReporteService usuarioService;
	@Autowired
	ParametroReporteService parametroReporteService;
	@Autowired
	private ReporteService reporteService;
	@Autowired
	private UsuarioService usuarioSeguridadService;
	@Autowired
	private FrenteAtencionService frenteAtencionService;
	@Autowired
	private GenericoService genericoService;
    /** Objeto ResourceBundle util */
    private ResourceBundle recursos;
	
	// Reporte Seguimiento
	private String seccionMes, seccionAnio, frenteAtencion;
	private List<SelectItem> listarSeccionesMes;
	private List<String> listarSeccionesAnio = new ArrayList<String>();
	Map<String,String> listaMeses;
	Map<String,String> listaAnios;
	//Map<String,String> listaFrenteAtencion;
	FechaUtil fechaUtil = new FechaUtil();
	String anioActual;
	private String inicializar;
	// Reporte Cierre de Mes
	private List<Reclamo> lstReporteCierreMes = new ArrayList<Reclamo>();
	private Reclamo reclamo = new Reclamo();
	private FrenteAtencion frenteAtencionBean = new FrenteAtencion();
	private FrenteAtencion frenteBean = new FrenteAtencion();
	private EquipoTrabajo equipoTrabajoBean = new EquipoTrabajo();
	private String nombreMes;
	private List<EquipoTrabajo> listaEquipoTrabajo = new ArrayList<EquipoTrabajo>();
	private List<FrenteAtencion> listaFrenteAtencion = new ArrayList<FrenteAtencion>();
	private Integer cantNroReclamosConsulta = 500;
	 
	public ReportesController(){
		listaFrenteAtencion.clear();
	}
	
	public String irReporteSeguimientoResultado(){
		
		String ruta = "";
		try {
			reclamo = new Reclamo();
			
			Date fechaActual = new Date();
			String dateString = fechaUtil.obtenerFechaStr(fechaActual);
			String anioTemp[]= dateString.split("/");
			anioActual=anioTemp[2];
			listaMeses=listarMeses();
			listaAnios=listarAnios();
			listaFrenteAtencion = obtenerListaFrenteAtencion();
			seccionAnio=anioActual;
			ruta = "reporteSeguimientoResultado?faces-redirect=true";
		} catch (Exception exception) {
			LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(exception)[1], exception);
	    	String[] error = MensajeExceptionUtil.obtenerMensajeError(exception);
	    	String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
	    	WebUtil.mensajeError(mensajeError);
		}
		return ruta;
	}
	

	public String irReporteCierredeMes(){
		
		  Date fechaLocal = new Date();
		  String dateString = fechaUtil.obtenerFechaStr(fechaLocal);
		  String anioTempTotal[]= dateString.split("/");
		  String anioTmp,mesTmp;
		  mesTmp=anioTempTotal[1];
		  nombreMes=fechaUtil.obtenerNombreMes(Integer.parseInt(mesTmp));
		  
		String ruta = "";
		try {
			listaMeses=listarMeses();
			listaAnios=listarAnios();
			
			ruta = "reporteCierreMes?faces-redirect=true";
		} catch (Exception exception) {
			LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(exception)[1], exception);
	    	String[] error = MensajeExceptionUtil.obtenerMensajeError(exception);
	    	String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
	    	WebUtil.mensajeError(mensajeError);
		}
		return ruta;
	}
	
	public  <K, V extends Comparable<? super V>> Map<K, V> 
    sortByValue( Map<K, V> map )
{
    List<Map.Entry<K, V>> list =
        new LinkedList<>( map.entrySet() );
    Collections.sort( list, new Comparator<Map.Entry<K, V>>()
    {
        public int compare( Map.Entry<K, V> o1, Map.Entry<K, V> o2 )
        {
            return (o1.getValue()).compareTo( o2.getValue() );
        }
    } );

    Map<K, V> result = new LinkedHashMap<>();
    for (Map.Entry<K, V> entry : list)
    {
        result.put( entry.getKey(), entry.getValue() );
    }
    return result;
}
	
	
    public  <K extends Comparable,V extends Comparable> Map<K,V> sortByKeys(Map<K,V> map){
        List<K> keys = new LinkedList<K>(map.keySet());
        Collections.sort(keys);
        Map<K,V> sortedMap = new LinkedHashMap<K,V>();
        for(K key: keys){
            sortedMap.put(key, map.get(key));
        }
      
        return sortedMap;
    }

	  public  Map<String,String> listarMeses(){
		  Map<String,String> listaMeses = new HashMap<String, String>();
		  listaMeses.put("Enero", 	  "01");
		  listaMeses.put("Febrero",   "02");
		  listaMeses.put("Marzo", 	  "03");
		  listaMeses.put("Abril", 	  "04");
		  listaMeses.put("Mayo", 	  "05");
		  listaMeses.put("Junio", 	  "06");
		  listaMeses.put("Julio", 	  "07");
		  listaMeses.put("Agosto", 	  "08");
		  listaMeses.put("Setiembre", "09");
		  listaMeses.put("Octubre",   "10");
		  listaMeses.put("Noviembre", "11");
		  listaMeses.put("Diciembre", "12");
		  return  sortByValue(listaMeses);
		  
	  }
	  
	  public  Map<String,String> listarAnios(){
		  Map<String,String> listaAnios = new HashMap<String, String>();
		 int contador=1;
		  for (int i = 2000; i <= 2025; i++) {
			  listaAnios.put(""+i, ""+i);
			  contador++;
		}
		  return  sortByKeys(listaAnios);
	  }
	  
	    
	  public void generarReporteCierreMes(){
		  try {
			  String cantidad = PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_ALFRESCO, "cantidad_numeros_reclamo_consulta");
			  cantNroReclamosConsulta = Integer.valueOf(cantidad);
			  	Empleado usuarioSeguridad = (Empleado) WebUtil.obtenerObjetoSesion(Constantes.SESION_USUARIOINICIO);
				List<ParametroReporte> listaParametrosReporte = new ArrayList<ParametroReporte>();
				  Session sessionCmis = (Session)WebUtil.obtenerObjetoSesion(Constantes.SESION_SESIONCMIS);
				  String where, where02;
				  List<Map<String, Object>> listaMetadataReporte = new ArrayList<Map<String,Object>>();
				  List<Map<String, Object>> listaMetadataReporteNoExistentes = new ArrayList<Map<String,Object>>();
				  List<Map<String, Object>> listaMeses = new ArrayList<Map<String,Object>>();
				  List<Map<String, Object>> listaMetadataReporteErrores = new ArrayList<Map<String,Object>>();
				  List<Map<String, Object>> listaMetadataRepErrores = new ArrayList<Map<String,Object>>();
				  Map<String,Object> mapReporteSeguimiento = new HashMap<String,Object>();
				  /*FECHA LOCAL */
				  Date fechaLocal = new Date();
				  String dateString = fechaUtil.obtenerFechaStr(fechaLocal);
				  String anioTempTotal[]= dateString.split("/");
				  String anioTmp,mesTmp;

				  /*REQ001 - Item001 - Inicio*/
				  /*mesTmp=anioTempTotal[1];
				  anioTmp=anioTempTotal[2];
				  */
				  mesTmp=seccionMes;
				  anioTmp=seccionAnio;

				  List<Map<String, Object>> listaMetadataReporteTmp = new ArrayList<Map<String,Object>>();
				  
				  /*Lista de reclamos que existen en el alfresco*/
				  List<Map<String, Object>> listaReclamosExisten = new ArrayList<Map<String,Object>>();
				  
				  
				  List<String> listaGeneralNumeroReclamos = new ArrayList<String>();
				  List<String> listaParametrosAux = new ArrayList<String>();
				  mapReporteSeguimiento.put("mes", seccionMes);
				  mapReporteSeguimiento.put("anio", seccionAnio);
				  List<List<String>> listaNumeroReclamos;
				  
					 if (parametroReporteService.validarParametroReporte(mapReporteSeguimiento)==0) {

					  listaParametrosReporte = parametroReporteService.listarParametroReporteSeguimiento(mapReporteSeguimiento);
				  	  if (listaParametrosReporte!=null && !listaParametrosReporte.isEmpty()){
				  		  if (validarParametroInicialFinal(listaParametrosReporte)) {
					  			listaGeneralNumeroReclamos = generarListaReporte(listaParametrosReporte);
								for(ParametroReporte valoresReporte: listaParametrosReporte ){
									listaNumeroReclamos = new ArrayList<List<String>>();
									listaNumeroReclamos = obtenerListaNumerosReclamo(valoresReporte);
									for (List<String> numeroReclamosConsulta : listaNumeroReclamos) {
										where = WebUtil.obtenerReporteWhereCmis(numeroReclamosConsulta);
									  	listaMetadataReporteTmp = alfrescoService.listarObjetosFiltro(sessionCmis, TablasQueryEnumSisRec.BBVAFOLDER.toString(), where);
									  	if (listaMetadataReporteTmp != null && !listaMetadataReporteTmp.isEmpty()) {
									  		listaReclamosExisten.addAll(listaMetadataReporteTmp);
										}
									}
								  }
							}else {
								WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_ValoresInicialFinal"));
								return;
							}
						  
						  // Recuperar files no existentes
						  listaMetadataReporteNoExistentes = obtenerListaReclamosNoExisten(listaGeneralNumeroReclamos,listaReclamosExisten); 


						  // Recuperar files Errores
						  	Map<String, Object> properties = new HashMap<String, Object>() ;
							String inicioMes=  anioTmp + "-"+ mesTmp + "-01 00:00:00";
							String finMes=  anioTmp + "-"+ mesTmp + "-31 00:00:00";
							properties.put("inicioMes", inicioMes);
							properties.put("finMes", finMes);
							listaMeses.add(properties);
						  	where = inicioMes;
						  	where02=finMes;
						  	listaMetadataReporteErrores= alfrescoService.listarObjetosReporteCierreMes(sessionCmis, TablasQueryEnumSisRec.BBVAFOLDER.toString(), where, where02);
						  	 Map<String, Object> prop = new HashMap<String, Object>() ;
						  	 for (int i = 0; i < listaMetadataReporteErrores.size(); i++) {
						  		  boolean encuentra = false;
								  for (int j = 0; j <listaReclamosExisten.size(); j++) {
									  prop = new HashMap<>();
									  if (listaMetadataReporteErrores.get(i).get("Nro Reclamo").toString().equals(listaReclamosExisten.get(j).get("Nro Reclamo").toString())) {
										  encuentra = true;
										  break;
									  }
									  
								  }
								  if(!encuentra){
									  listaMetadataRepErrores.add(listaMetadataReporteErrores.get(i));
								  }
						  	 }
						  	 
						  	Map<String, Object> reporteCierreMes = new HashMap<String, Object>() ;
						  	reporteCierreMes.put("idUsuaCrea", WebUtil.obtenerLoginUsuario());
						  	reporteCierreMes.put("deTerminalCrea", WebUtil.obtenerTerminal());
						  	reporteCierreMes.put("listaReclamosExistentes", listaReclamosExisten);
						  	reporteCierreMes.put("listaReclamosNoExistentes", listaMetadataReporteNoExistentes);
						  	reporteCierreMes.put("listaReclamosErrores", listaMetadataRepErrores);
						  	reporteCierreMes.put("mes", seccionMes);
							reporteCierreMes.put("anio", seccionAnio);
							/*REGISTRO EN BASE DE DATOS */
						  	reporteService.registrarReporteCierreMes(reporteCierreMes);
							
				  	}else{
						 WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_listaVacia"));
					}
				  } else { 
					  List<Map<String, Object>> listaReclamosCierreMesBaseDatos = reporteService.listarReclamosCierreMes(mapReporteSeguimiento);
					  listaReclamosExisten = obtenerReclamosBaseDatos(listaReclamosCierreMesBaseDatos,Constantes.ID_ESTADO_RECLAMO_FILE_ABIERTO ,Constantes.ID_ESTADO_RECLAMO_FILE_CERRADO );
					  listaMetadataReporteNoExistentes = obtenerReclamosBaseDatos(listaReclamosCierreMesBaseDatos,Constantes.ID_ESTADO_RECLAMO_FILE_NO_EXISTENTE);
					  listaMetadataRepErrores = obtenerReclamosBaseDatos(listaReclamosCierreMesBaseDatos,Constantes.ID_ESTADO_RECLAMO_FILE_ERROR);
				  }
						  // Reporte Excel
						  LOGGER.info("Inicio generarReporteExcel");
						  
						  HttpServletResponse response = (HttpServletResponse) FacesContext.getCurrentInstance().getExternalContext().getResponse();
						  InputStream flujoBytesExcel=null;
						  flujoBytesExcel = getClass().getResourceAsStream("/plantillaReporteExcel.xls");
						  HSSFWorkbook libroExcel = new HSSFWorkbook(flujoBytesExcel);
						  HSSFSheet hojaExcel = libroExcel.getSheetAt(0);
						  Integer totalRegistros  = (listaReclamosExisten.size() + listaMetadataReporteNoExistentes.size() + listaMetadataRepErrores.size());
						  String periodo = ( fechaUtil.obtenerNombreMes(Integer.valueOf(mesTmp) ) )+ " - " + anioTmp;
						  
						  UDocumentoExcel uDocumentoExcel = new UDocumentoExcel(libroExcel);
						  HSSFCellStyle csCeldaFilaBlanca = uDocumentoExcel.generarEstiloCeldaTablaBlanco(HSSFCellStyle.ALIGN_LEFT);
						  HSSFCellStyle csCeldaFilaColor = uDocumentoExcel.generarEstiloCeldaTablaColor(HSSFCellStyle.ALIGN_LEFT);
						  
						  HSSFRow filaExcel = null;
						  HSSFCell celdaExcel = null;
						  
						  /* Ancho por defecto de las columnas; importante no se configura el ancho de la celda 0 ya que esta viene de la plantilla excel por defecto. */
						  hojaExcel.setColumnWidth((short)(1),(short)4000);
				          hojaExcel.setColumnWidth((short)(2),(short)5000); 
				          hojaExcel.setColumnWidth((short)(3),(short)4000);
				          hojaExcel.setColumnWidth((short)(4),(short)8000);
				         
				          
				          filaExcel = hojaExcel.createRow(4);
				          celdaExcel = filaExcel.createCell((short)0);

				          //Insertar el total de registros 
				          filaExcel = hojaExcel.createRow(5);
				          celdaExcel = filaExcel.createCell((short)0);
				          celdaExcel.setCellValue("Total Registros encontrados = " + (listaReclamosExisten.size() + listaMetadataReporteNoExistentes.size() + listaMetadataRepErrores.size()));
				          //Insertar el mes y año del reporte
				          celdaExcel = filaExcel.createCell((short)4);
				          celdaExcel.setCellValue("Periodo: " +( fechaUtil.obtenerNombreMes(Integer.valueOf(mesTmp) ) )+ " - " + anioTmp);

				          
				          //Insertar en nombre del reporte
				          filaExcel = hojaExcel.getRow(2);
				          celdaExcel = filaExcel.getCell((short)3);
				          celdaExcel.setCellValue("REPORTE - CIERRE DE MES");
				          
				
			        	  //Insertar la fecha
				          filaExcel = hojaExcel.getRow(1);
				          celdaExcel = filaExcel.getCell((short)7);
				          Calendar fechaActual = Calendar.getInstance();
				          String strFechaActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.DAY_OF_MONTH)+"",2)+"/"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MONTH)+1)+"",2)+"/"+Funciones.CompletaCerosIzq(fechaActual.get(Calendar.YEAR)+"",4);
				          celdaExcel.setCellValue(strFechaActual);
						  
				          // Insertar la Hora
				          filaExcel = hojaExcel.getRow(2);
				          celdaExcel = filaExcel.getCell((short)7);
				          String sHoraActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.HOUR)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MINUTE)+1)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.SECOND)+1)+"",2);
				          celdaExcel.setCellValue(sHoraActual);
				          
				          // Insertar el Usuario
				          filaExcel = hojaExcel.getRow(3);
				          celdaExcel = filaExcel.getCell((short)7);    
				          celdaExcel.setCellValue( usuarioSeguridad.getLogin());
//				          
				          /* Fila desde la que se inicia. */
				          short numeroFila = 7;
				          int contaReporte=1;
				          // Datos de la tabla detalle:
				          short numeroFilaTotal = 0;
				          if(listaReclamosExisten!=null && !listaReclamosExisten.isEmpty()){
				          	 String strTemporal = "";  
				          	 int contador=0;
				          	int  mes=0;
				          	String fecha="", codUsuarioRegistro,  nombreCompleto="";
				        	Map<String,Object> parametersUsuario = new HashMap<String,Object>();   
				        	UsuarioReporte usuario = new UsuarioReporte();
				          	for (Map<String, Object>  reporteSeguimiento : listaReclamosExisten ) {
				          		numeroFilaTotal ++;
				          		if(numeroFilaTotal > Constantes.NUMERO_MAXIMO_FILAS_EXCEL){
				          			numeroFilaTotal = 0;
				          			hojaExcel=crearHojaExcelReporteCierreMes(libroExcel, "Reporte Cierre Mes ", periodo, totalRegistros,"REPORTE - CIERRE DE MES");
				          			numeroFila = 8;
								}
				          		
				          	    //Inicio traer nombre Usuario
				          		if (!reporteSeguimiento.get("Cod Registro").toString().equals("") && reporteSeguimiento.get("Cod Registro").toString()!=null){
				          			usuario.setLogin(reporteSeguimiento.get("Cod Registro").toString());
				          			usuario =usuarioService.nombreUsuario(usuario);
				          			usuario.setNombreCompleto(usuario.getNombre() + " " + usuario.getApellidoPaterno() + " " + usuario.getApellidoMaterno() );
				          		}
				          	    //Fin traer nombre Usuario
				          		
				          		/* Estilo de la fila. */
				          		HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
				          		
				          		filaExcel = hojaExcel.createRow(++numeroFila);
				                celdaExcel = filaExcel.createCell((short)0);
				                strTemporal = String.valueOf( contaReporte );
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				          		
				                celdaExcel = filaExcel.createCell((short)1);
				                strTemporal =  reporteSeguimiento.get("Nro Reclamo").toString();
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                
				          
				                celdaExcel = filaExcel.createCell((short)2);
				                strTemporal =  reporteSeguimiento.get("Estado").toString();
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				              
				                celdaExcel = filaExcel.createCell((short)3);
				                strTemporal = reporteSeguimiento.get("Frente").toString();
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                  
				                celdaExcel = filaExcel.createCell((short)4);
				                strTemporal = reporteSeguimiento.get("Equipo").toString();
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				               
				                celdaExcel = filaExcel.createCell((short)5); 
				                strTemporal =  usuario.getNombreCompleto();
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				               
				                celdaExcel = filaExcel.createCell((short)6);
				                celdaExcel.setCellStyle(estilo);
				                
				                if (reporteSeguimiento.get("Fecha Registro Reclamo")!=null && !reporteSeguimiento.get("Fecha Registro Reclamo").equals("") ){
				                	 if ( reporteSeguimiento.get("Fecha Registro Reclamo").toString()!=null && !reporteSeguimiento.get("Fecha Registro Reclamo").toString().equals("") ) {
						                	fecha= reporteSeguimiento.get("Fecha Registro Reclamo").toString();
							          		String fechaTemp[]=  fecha.split("/");
							          		mes=Integer.valueOf(fechaTemp[1]);
							                strTemporal =  fechaUtil.obtenerNombreMes(mes);
							                celdaExcel.setCellValue(strTemporal);
							                celdaExcel.setCellStyle(estilo);
										}
				                }
				               
				                if (reporteSeguimiento.get("Fecha Registro Reclamo")!=null && !reporteSeguimiento.get("Fecha Registro Reclamo").equals("")){
				                celdaExcel = filaExcel.createCell((short)7);
				                strTemporal =  reporteSeguimiento.get("Fecha Registro Reclamo").toString();
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                }
				                
				                celdaExcel = filaExcel.createCell((short)7);
				                strTemporal = "";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                
				                /*REQ001 - Item001 - Inicio*/
				                if (reporteSeguimiento.get("Fecha Modificación")!=null && !reporteSeguimiento.get("Fecha Modificación").equals("")){
				                celdaExcel = filaExcel.createCell((short)8);
				                 strTemporal =  reporteSeguimiento.get("Fecha Modificación").toString();
				                 celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                }
				                celdaExcel = filaExcel.createCell((short)8);
				                 strTemporal = "";
				                 celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                /*REQ001 - Item001 - Fin*/
				                contaReporte++;
				          	}
				          }
				          
				          if(listaMetadataReporteNoExistentes!=null && !listaMetadataReporteNoExistentes.isEmpty()){
				        		for (Map<String, Object>  reporteNoExistentes : listaMetadataReporteNoExistentes ) {
				        			numeroFilaTotal ++;
				        			if(numeroFilaTotal > Constantes.NUMERO_MAXIMO_FILAS_EXCEL){
				        				numeroFilaTotal = 0;
					          			hojaExcel=crearHojaExcelReporteCierreMes(libroExcel, "Reporte Cierre Mes ", periodo, totalRegistros,"REPORTE - CIERRE DE MES");
					          			numeroFila = 8;
									}
				        			String strTemporal = "";  
				        			int contador=0;
				        			/* Estilo de la fila. */
					          		HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
					          		
					          		filaExcel = hojaExcel.createRow(++numeroFila);
					                celdaExcel = filaExcel.createCell((short)0);
					                strTemporal = String.valueOf( contaReporte );
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					          		
					                celdaExcel = filaExcel.createCell((short)1);
					                strTemporal =  reporteNoExistentes.get("Nro Reclamo").toString();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                celdaExcel = filaExcel.createCell((short)2);
					                strTemporal =  "NO EXISTE";
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                celdaExcel = filaExcel.createCell((short)3);
					                strTemporal = " - ";
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                celdaExcel = filaExcel.createCell((short)4);
					                strTemporal = " - ";
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                celdaExcel = filaExcel.createCell((short)5); 
					                strTemporal =  " - ";
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                celdaExcel = filaExcel.createCell((short)6);
					                strTemporal =   " - ";
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                celdaExcel = filaExcel.createCell((short)7);
					                strTemporal =   " - ";
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                /*REQ001 - Item001 - Inicio*/
					                celdaExcel = filaExcel.createCell((short)8);
					                strTemporal =   " - ";
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                /*REQ001 - Item001 - Fin*/
					                contaReporte++;
				        		}
				          }
				          if(listaMetadataRepErrores!=null && !listaMetadataRepErrores.isEmpty()){
				        		UsuarioReporte usuario = new UsuarioReporte();
				        		String fecha="";
				        		int  mes=0;
				        	  for (Map<String, Object>  reporteErrores : listaMetadataRepErrores ) {
				        		  numeroFilaTotal ++;
				        		  if(numeroFilaTotal > Constantes.NUMERO_MAXIMO_FILAS_EXCEL){
				        			  numeroFilaTotal = 0;
					          			hojaExcel=crearHojaExcelReporteCierreMes(libroExcel, "Reporte Cierre Mes ", periodo, totalRegistros,"REPORTE - CIERRE DE MES");
					          			numeroFila = 7;
								  }
				        			//Inicio obtener Mes 
				        		  if (!reporteErrores.get("Fecha Registro Reclamo").equals("") && reporteErrores.get("Fecha Registro Reclamo").toString()!=null){
					          		if (!reporteErrores.get("Fecha Registro Reclamo").toString().equals("") && reporteErrores.get("Fecha Registro Reclamo").toString()!=null){
					          			fecha= reporteErrores.get("Fecha Registro Reclamo").toString();
						          		String fechaTemp[]=  fecha.split("/");
						          		mes=Integer.valueOf(fechaTemp[1]);
					          		}
				        		  }
					          		//Fin obtener Mes 
					          		
				        		  if (!reporteErrores.get("Cod Registro").toString().equals("") && reporteErrores.get("Cod Registro").toString()!=null){
					          			usuario.setLogin(reporteErrores.get("Cod Registro").toString());
					          			usuario =usuarioService.nombreUsuario(usuario);
					          			if (usuario!=null && usuario.getNombre()!=null && !usuario.getNombre().equals("")){
					          				usuario.setNombreCompleto(usuario.getNombre() + " " + usuario.getApellidoPaterno() + " " + usuario.getApellidoMaterno() );
					          			}else{
					          				usuario=new UsuarioReporte();
					          				usuario.setNombreCompleto(" - ");
					          			}
					          		}
				        		  
				        		  String strTemporal = "";  
				        			int contador=0;
				        			/* Estilo de la fila. */
					          		HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
					          		
					          		filaExcel = hojaExcel.createRow(++numeroFila);
					                celdaExcel = filaExcel.createCell((short)0);
					                strTemporal = String.valueOf( contaReporte );
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					          		
					                celdaExcel = filaExcel.createCell((short)1);
					                strTemporal =  reporteErrores.get("Nro Reclamo").toString();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					          
					                celdaExcel = filaExcel.createCell((short)2);
					                strTemporal = "ERROR - " +  reporteErrores.get("Estado").toString();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					              
					                celdaExcel = filaExcel.createCell((short)3);
					                strTemporal = reporteErrores.get("Frente").toString();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                celdaExcel = filaExcel.createCell((short)4);
					                strTemporal = reporteErrores.get("Equipo").toString();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					               
					                celdaExcel = filaExcel.createCell((short)5); 
					                strTemporal =  usuario.getNombreCompleto();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					               
					                celdaExcel = filaExcel.createCell((short)6);
					                strTemporal =  fechaUtil.obtenerNombreMes(mes);
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                
					                if (!reporteErrores.get("Fecha Registro Reclamo").equals("") && reporteErrores.get("Fecha Registro Reclamo")!=null){
					                celdaExcel = filaExcel.createCell((short)7);
					                strTemporal =  reporteErrores.get("Fecha Registro Reclamo").toString();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                }
					                	    
					                	   celdaExcel = filaExcel.createCell((short)7);
							                strTemporal =  "";
							                celdaExcel.setCellValue(strTemporal);
							                celdaExcel.setCellStyle(estilo);
					                
					                
					                /*REQ001 - Item001 - Inicio*/
					                if ( reporteErrores.get("Fecha Modificación")!=null && !reporteErrores.get("Fecha Modificación").equals("") ){
					                celdaExcel = filaExcel.createCell((short)8);
					                strTemporal =  reporteErrores.get("Fecha Modificación").toString();
					                celdaExcel.setCellValue(strTemporal);
					                celdaExcel.setCellStyle(estilo);
					                }else{
					                	 celdaExcel = filaExcel.createCell((short)8);
							                strTemporal =  "";
							                celdaExcel.setCellValue(strTemporal);
							                celdaExcel.setCellStyle(estilo);
					                }
					                	
					                
					                /*REQ001 - Item001 - Fin*/
					                contaReporte++;
				        	  }
				        		
				          }
				          
				          /* Pie de pagina del reporte. */
				          numeroFila=(short)(numeroFila + 2);
				          filaExcel = hojaExcel.createRow(numeroFila);
				          celdaExcel = filaExcel.createCell((short)0);
				          celdaExcel.setCellStyle(uDocumentoExcel.generarEstiloLeyenda(HSSFCellStyle.ALIGN_CENTER, false));

				          /* Enviar archivo al navegador para su descarga. */
				          response.setContentType("application/vnd.ms-excel");
				          response.setHeader("Content-Disposition","attachment; filename=\"ReporteCierreMes.xls\"");
				          OutputStream out = response.getOutputStream();
				          libroExcel.write(out);
				          out.flush();
				          out.close();
				          FacesContext.getCurrentInstance().responseComplete();
/*
					}else{
						
						 WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_listaVacia"));
					}
					  	
	*/			 

		    	}catch(Exception exception){
		    		LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(exception)[1], exception);
		 	    	String[] error = MensajeExceptionUtil.obtenerMensajeError(exception);
		 	    	String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
		 	    	WebUtil.mensajeError(mensajeError);
		        }
			  LOGGER.info("Fin generarReporteExcel"); 
		  
		  
	  }
	  /*
	  /*REQ001 - Item001 - Inicio*/
	  public void generarReporteCierreMesResumen() throws GmdException {
		  try {
			  Empleado usuarioSeguridad = (Empleado) WebUtil.obtenerObjetoSesion(Constantes.SESION_USUARIOINICIO);
			  String anioActualReporte;
			  List<Reclamo> listaParaMetrosInterv = new ArrayList<Reclamo>();
			  int anioPasado;
			  Date fechaActualReporte = new Date();
			  String dateStringReporte = fechaUtil.obtenerFechaStr(fechaActualReporte);
			  String anioTempReporte[]= dateStringReporte.split("/");
			  anioActualReporte=anioTempReporte[2];
			  String anioTmp,mesTmp;
			  mesTmp=seccionMes;
			  anioTmp=seccionAnio;
			  Map<String,Object> mapReporteResumenCierreMes = new HashMap<String,Object>();
			  mapReporteResumenCierreMes.put("mes", seccionMes);
			  mapReporteResumenCierreMes.put("anio", seccionAnio);
			 
			  
			  if (!seccionMes.equals("")&& !seccionAnio.equals("")){
				  listaParaMetrosInterv = reporteService.listarReclamosAnioActual(mapReporteResumenCierreMes);
				  if (listaParaMetrosInterv!=null && !listaParaMetrosInterv.isEmpty() ){
				  List<Reclamo> listaReclamosHistorico = new ArrayList<Reclamo>();
				  List<Reclamo> listaReclamosMes = new ArrayList<Reclamo>();
				  Map<String, Object> listaMatriz = new HashMap<String, Object>();
				  int  totalAnioActualCerrado=0, totalAnioActualAbierto=0, totalAnioActualNoExiste=0, totalAnioActualError=0;
				  Map<String, Object> listaAnioActualMes = new HashMap<String,Object>();
				  Reclamo reclamo= new Reclamo();
	 
				  List<Reclamo> listaReporteAnioPasado = obtenerListaReclamosHistorico(listaReclamosHistorico);

				  int contaMes=0;
				  String descMes;List<String> listaMes = new ArrayList<String>();
				  reclamo.setAnio(seccionAnio);
				  
				    
				  reclamo.setMes(seccionMes);
				  descMes=fechaUtil.obtenerNombreMes(Integer.parseInt(seccionMes));
				  listaReclamosMes=reporteService.listarDatosAnioActual(reclamo);
				  if (listaReclamosMes!=null && listaReclamosMes.size()>0){
					  //INICIO
					  List<Reclamo> listaAnioActual = obtenerListaReclamosHistorico(listaReclamosMes);
					  //FIN
					  listaMatriz.put(descMes,listaAnioActual );
					  listaMes.add(descMes);
				  }
		 
				  
	 
				 int cantTotalMes01=0, cantTotalMes02=0,cantTotalMes03=0,cantTotalMes04=0,cantTotalMes05=0;
				  int cantTotalAnio01=0,cantTotalAnio02=0,cantTotalAnio03=0,cantTotalAnio04=0,cantTotalAnio05=0;
				  
				  //Fin de obtener la lista de los meses del anio actual
				  
				  // Reporte Excel
				  HttpServletResponse response = (HttpServletResponse) FacesContext.getCurrentInstance().getExternalContext().getResponse();
				  InputStream flujoBytesExcel=null;
				  flujoBytesExcel = getClass().getResourceAsStream("/plantillaReporte.xls");
				  HSSFWorkbook libroExcel = new HSSFWorkbook(flujoBytesExcel);
				  HSSFSheet hojaExcel = libroExcel.getSheetAt(0);
				  
				  UDocumentoExcel uDocumentoExcel = new UDocumentoExcel(libroExcel);
				  HSSFCellStyle csCeldaFilaBlanca = uDocumentoExcel.generarEstiloCeldaTablaBlanco(HSSFCellStyle.ALIGN_LEFT);
				  HSSFCellStyle csCeldaFilaColor = uDocumentoExcel.generarEstiloCeldaTablaColor(HSSFCellStyle.ALIGN_LEFT);
				  HSSFCellStyle estiloCabeceraTabla = uDocumentoExcel.generarEstiloCabeceraTabla();
				  
				  HSSFRow filaExcel = null;
				  HSSFCell celdaExcel = null;
				  
				  /* Ancho por defecto de las columnas; importante no se configura el ancho de la celda 0 ya que esta viene de la plantilla excel por defecto. */
				  hojaExcel.setColumnWidth((short)(1),(short)4000);
				  hojaExcel.setColumnWidth((short)(2),(short)4000); 
				  hojaExcel.setColumnWidth((short)(3),(short)4000);
				  hojaExcel.setColumnWidth((short)(4),(short)8000);
				  
				  filaExcel = hojaExcel.createRow(4);
				  celdaExcel = filaExcel.createCell((short)0);
				  		 
				  //Insertar en nombre del reporte
				  filaExcel = hojaExcel.getRow(2);
				  celdaExcel = filaExcel.getCell((short)2);
				  celdaExcel.setCellValue("REPORTE RESUMEN CIERRE MES");
				  //Insertar la fecha
				  filaExcel = hojaExcel.getRow(1);
				  celdaExcel = filaExcel.getCell((short)6);
				  Calendar fechaActual = Calendar.getInstance();
				  String strFechaActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.DAY_OF_MONTH)+"",2)+"/"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MONTH)+1)+"",2)+"/"+Funciones.CompletaCerosIzq(fechaActual.get(Calendar.YEAR)+"",4);
				  celdaExcel.setCellValue(strFechaActual);
				  
				  // Insertar la Hora
				  filaExcel = hojaExcel.getRow(2);
				  celdaExcel = filaExcel.getCell((short)6);
				  String sHoraActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.HOUR)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MINUTE)+1)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.SECOND)+1)+"",2);
				  celdaExcel.setCellValue(sHoraActual);
				  
				  // Insertar el Usuario
				  filaExcel = hojaExcel.getRow(3);
				  celdaExcel = filaExcel.getCell((short)6);
				  celdaExcel.setCellValue( usuarioSeguridad.getLogin());
				  
				   //Insertar el mes y año del reporte
				  filaExcel = hojaExcel.getRow(5);
		          celdaExcel = filaExcel.createCell((short)3);
		          celdaExcel.setCellValue("Periodo: " +( fechaUtil.obtenerNombreMes(Integer.valueOf(mesTmp) ) )+ " - " + anioTmp);
		    
				  
				  

				  int numTemp = 1;
				  for (String string : listaMes) {
					  filaExcel = hojaExcel.getRow(7);
					  celdaExcel = filaExcel.getCell((short)numTemp);
					  celdaExcel.setCellValue(string + " " + String.valueOf(anioActualReporte));
					  celdaExcel.setCellStyle(estiloCabeceraTabla);
					  numTemp ++;
				  }

				
				  int totalReporte01=0,totalReporte02=0,totalReporte03=0,totalReporte04=0,totalReporte05=0,totalReporte06=0,totalReporte07=0,totalReporte08=0,totalReporte09=0,totalReporte10=0,totalReporte11=0,totalReporte12=0 ;
				  short numeroFila = 7;
				  
				  // Datos de la tabla detalle:
				  if(listaMatriz!=null && listaMatriz.size()>0){
					  /***INICIO DETALLE - EXCEL**/
					  int contador1=0;
					  int i=0;
					  for (Reclamo objReporteAnioPasadoj :  listaReporteAnioPasado){
						  String strTemporal = "";  
						  HSSFCellStyle estilo = ((contador1%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
					  
						  filaExcel = hojaExcel.createRow(++numeroFila);
						  int num = 1;
						  if (i == 0) { // FILE ABIERTO
							  celdaExcel = filaExcel.createCell((short)0);
						      celdaExcel.setCellValue("FILE ABIERTO");
						      celdaExcel.setCellStyle(estilo);
						      

							  for (String string : listaMes) {
								  cantTotalMes02 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_ABIERTO);
								  celdaExcel = filaExcel.createCell((short)num);
								  strTemporal = String.valueOf(cantTotalMes02);
								  totalAnioActualAbierto+=cantTotalMes02;
								  celdaExcel.setCellValue(strTemporal);
								  celdaExcel.setCellStyle(estilo);
								  num ++;
								  if(string.equals("ENERO")){
						        	totalReporte01=totalReporte01+cantTotalMes02;
								  }else if (string.equals("FEBRERO")){
						        	totalReporte02=totalReporte02+cantTotalMes02;
								  }else if (string.equals("MARZO")){
						        	totalReporte03=totalReporte03+cantTotalMes02;
								  }else if (string.equals("ABRIL")){
						        	totalReporte04=totalReporte04+cantTotalMes02;
								  }else if (string.equals("MAYO")){
						        	totalReporte05=totalReporte05+cantTotalMes02;
								  }else if (string.equals("JUNIO")){
						        	totalReporte06=totalReporte06+cantTotalMes02;
								  }else if (string.equals("JULIO")){
						        	totalReporte07=totalReporte07+cantTotalMes02;
								  }else if (string.equals("AGOSTO")){
						        	totalReporte08=totalReporte08+cantTotalMes02;
								  }else if (string.equals("SETIEMBRE")){
						        	totalReporte09=totalReporte09+cantTotalMes02;
								  }else if (string.equals("OCTUBRE")){
						        	totalReporte10=totalReporte10+cantTotalMes02;
								  }else if (string.equals("NOVIEMBRE")){
						        	totalReporte11=totalReporte11+cantTotalMes02;
								  }else if (string.equals("DICIEMBRE")){
						        	totalReporte12=totalReporte12+cantTotalMes02;
								  }
							  }

						  }else if (i == 1) { // FILE CERRADO
							  celdaExcel = filaExcel.createCell((short)0);
						      celdaExcel.setCellValue("FILE CERRADO");
						      celdaExcel.setCellStyle(estilo);
							  celdaExcel = filaExcel.createCell((short)1);
				 
						      celdaExcel.setCellValue(cantTotalAnio01);
						      celdaExcel.setCellStyle(estilo);
					      
							  for (String string : listaMes) {
								  cantTotalMes01 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_CERRADO);
								  celdaExcel = filaExcel.createCell((short)num);
								  strTemporal = String.valueOf(cantTotalMes01);///27
								  totalAnioActualCerrado+=cantTotalMes01;
								  celdaExcel.setCellValue(strTemporal);
								  celdaExcel.setCellStyle(estilo);
								  num ++;
								  if(string.equals("ENERO")){
									  totalReporte01=totalReporte01+cantTotalMes01;
								  }else if (string.equals("FEBRERO")){
									  totalReporte02=totalReporte02+cantTotalMes01;
								  }else if (string.equals("MARZO")){
						        	totalReporte03=totalReporte03+cantTotalMes01;
								  }else if (string.equals("ABRIL")){
									  totalReporte04=totalReporte04+cantTotalMes01;
								  }else if (string.equals("MAYO")){
									  totalReporte05=totalReporte05+cantTotalMes01;
								  }else if (string.equals("JUNIO")){
									  totalReporte06=totalReporte06+cantTotalMes01;
								  }else if (string.equals("JULIO")){
									  totalReporte07=totalReporte07+cantTotalMes01;
								  }else if (string.equals("AGOSTO")){
									  totalReporte08=totalReporte08+cantTotalMes01;
								  }else if (string.equals("SETIEMBRE")){
									  totalReporte09=totalReporte09+cantTotalMes01;
								  }else if (string.equals("OCTUBRE")){
									  totalReporte10=totalReporte10+cantTotalMes01;
								  }else if (string.equals("NOVIEMBRE")){
									  totalReporte11=totalReporte11+cantTotalMes01;
								  }else if (string.equals("DICIEMBRE")){
									  totalReporte12=totalReporte12+cantTotalMes01;
								  }
							  }

						  }else if (i == 2) { // ERROR
							  celdaExcel = filaExcel.createCell((short)0);
						      celdaExcel.setCellValue("ERROR");
						      celdaExcel.setCellStyle(estilo);
							  celdaExcel = filaExcel.createCell((short)1);

						      celdaExcel.setCellValue(cantTotalAnio04);
						      celdaExcel.setCellStyle(estilo);
							  for (String string : listaMes) {
								  cantTotalMes04 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_ERROR);
								  celdaExcel = filaExcel.createCell((short)num);
						          strTemporal = String.valueOf(cantTotalMes04);
						          totalAnioActualError+=cantTotalMes04;
						          celdaExcel.setCellValue(strTemporal);
						          celdaExcel.setCellStyle(estilo);
						          num ++;
						          if(string.equals("ENERO")){
						        	totalReporte01=totalReporte01+cantTotalMes04;
						          }else if (string.equals("FEBRERO")){
						        	totalReporte02=totalReporte02+cantTotalMes04;
						          }else if (string.equals("MARZO")){
						        	totalReporte03=totalReporte03+cantTotalMes04;
						          }else if (string.equals("ABRIL")){
						        	totalReporte04=totalReporte04+cantTotalMes04;
						          }else if (string.equals("MAYO")){
						        	totalReporte05=totalReporte05+cantTotalMes04;
						          }else if (string.equals("JUNIO")){
						        	totalReporte06=totalReporte06+cantTotalMes04;
						          }else if (string.equals("JULIO")){
						        	totalReporte07=totalReporte07+cantTotalMes04;
						          }else if (string.equals("AGOSTO")){
						        	totalReporte08=totalReporte08+cantTotalMes04;
						          }else if (string.equals("SETIEMBRE")){
						        	totalReporte09=totalReporte09+cantTotalMes04;
						          }else if (string.equals("OCTUBRE")){
						        	totalReporte10=totalReporte10+cantTotalMes04;
						          }else if (string.equals("NOVIEMBRE")){
						        	totalReporte11=totalReporte11+cantTotalMes04;
						          }else if (string.equals("DICIEMBRE")){
						        	totalReporte12=totalReporte12+cantTotalMes04;
						          }
							  }
						  }else if (i == 3) {// NO EXISTE
							  celdaExcel = filaExcel.createCell((short)0);
						      celdaExcel.setCellValue("NO EXISTE");
						      celdaExcel.setCellStyle(estilo);
							  celdaExcel = filaExcel.createCell((short)1);
						      celdaExcel.setCellValue(cantTotalAnio03);
						      celdaExcel.setCellStyle(estilo);
							  for (String string : listaMes) {
								  cantTotalMes03 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_NO_EXISTENTE);
								  celdaExcel = filaExcel.createCell((short)num);
								  strTemporal = String.valueOf(cantTotalMes03);
								  totalAnioActualNoExiste+=cantTotalMes03;
								  celdaExcel.setCellValue(strTemporal);
								  celdaExcel.setCellStyle(estilo);
						          num ++;
						          if(string.equals("ENERO")){
						        	totalReporte01=totalReporte01+cantTotalMes03;
						          }else if (string.equals("FEBRERO")){
						        	totalReporte02=totalReporte02+cantTotalMes03;
						          }else if (string.equals("MARZO")){
						        	totalReporte03=totalReporte03+cantTotalMes03;
						          }else if (string.equals("ABRIL")){
						        	totalReporte04=totalReporte04+cantTotalMes03;
						          }else if (string.equals("MAYO")){
						        	totalReporte05=totalReporte05+cantTotalMes03;
						          }else if (string.equals("JUNIO")){
						        	totalReporte06=totalReporte06+cantTotalMes03;
						          }else if (string.equals("JULIO")){
						        	totalReporte07=totalReporte07+cantTotalMes03;
						          }else if (string.equals("AGOSTO")){
						        	totalReporte08=totalReporte08+cantTotalMes03;
						          }else if (string.equals("SETIEMBRE")){
						        	totalReporte09=totalReporte09+cantTotalMes03;
						          }else if (string.equals("OCTUBRE")){
						        	totalReporte10=totalReporte10+cantTotalMes03;
						          }else if (string.equals("NOVIEMBRE")){
						        	totalReporte11=totalReporte11+cantTotalMes03;
						          }else if (string.equals("DICIEMBRE")){
						            	totalReporte12=totalReporte12+cantTotalMes03;
						          }
							  }
						  }
						  i++;
					  }
					  /***FIN DETALLE - EXCEL**/
					  filaExcel = hojaExcel.createRow(++numeroFila);
					  int contador=0;
					  int totalGeneralReporte=0;
					  totalGeneralReporte=totalAnioActualCerrado+totalAnioActualAbierto+totalAnioActualNoExiste+totalAnioActualError;
					  HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
					  
					  String strTemporalTotal = "";  
					  short contadorCelda=0;
					  celdaExcel = filaExcel.createCell(contadorCelda);
					  strTemporalTotal = "TOTAL";
			         celdaExcel.setCellValue(strTemporalTotal);
			          celdaExcel.setCellStyle(estilo);
			          contadorCelda++;
		          
			          celdaExcel = filaExcel.createCell(contadorCelda);
			          celdaExcel.setCellValue(strTemporalTotal);
			          celdaExcel.setCellStyle(estilo);
			          contadorCelda++;
	 		          contadorCelda--;
			          celdaExcel = filaExcel.createCell(contadorCelda);
			          strTemporalTotal = String.valueOf(totalGeneralReporte);
			          celdaExcel.setCellValue(strTemporalTotal);	
			          celdaExcel.setCellStyle(estilo);
				  }
				  /* Pie de pagina del reporte. */
				  numeroFila=(short)(numeroFila + 1);
				  filaExcel = hojaExcel.createRow(numeroFila);
				  celdaExcel = filaExcel.createCell((short)0);
				  celdaExcel.setCellStyle(uDocumentoExcel.generarEstiloLeyenda(HSSFCellStyle.ALIGN_CENTER, false));
				  celdaExcel.setCellValue("****FIN DE REPORTE****");
				  
				  /* Enviar archivo al navegador para su descarga. */
				  response.setContentType("application/vnd.ms-excel");
				  response.setHeader("Content-Disposition","attachment; filename=\"ReporteResumenCierreMes.xls\"");
				  OutputStream out = response.getOutputStream();
				  libroExcel.write(out);
				  out.flush();
				  out.close();
				  
				  }else{
					  WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_listaReclamosVacia"));
				  }
				}else{
				  WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_seleccionAnioMes"));
				}
			
		  } catch (Exception exception) {
			LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(exception)[1], exception);
	    	String[] error = MensajeExceptionUtil.obtenerMensajeError(exception);
	    	String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
	    	WebUtil.mensajeError(mensajeError);
		  }
	  }
	  /*REQ001 - Item001 - Fin*/
	  
    private List<Map<String, Object>> obtenerListaReclamosNoExisten( List<String> listaGeneralNumeroReclamos, List<Map<String, Object>> listaReclamosExisten) throws GmdException {
    	List<Map<String, Object>> listaMetadataReporteNoExistentes = new ArrayList<Map<String,Object>>();
		try {
			  Map<String, Object> propiedades;
			  for (int i = 0; i < listaGeneralNumeroReclamos.size(); i++) {
				  boolean encuentra = false;
				  for (int j = 0; j <listaReclamosExisten.size(); j++) {
					  if (listaReclamosExisten.get(j).get("Nro Reclamo").toString().equals(listaGeneralNumeroReclamos.get(i))) {
						  encuentra = true;
						  break;
					  }
				  	}
				  if(!encuentra){
					  propiedades = new HashMap<String, Object>() ;
					  String columna = "Nro Reclamo";
					  Object valor = listaGeneralNumeroReclamos.get(i);
					  propiedades.put(columna, valor);
					  listaMetadataReporteNoExistentes.add(propiedades);
				  }	 
			  }
		} catch (Exception excepcion) {
			throw new GmdException(excepcion);
		}
		return listaMetadataReporteNoExistentes;
	}
	private boolean validarParametroInicialFinal(List<ParametroReporte> listaParametrosReporte) throws GmdException {
    	boolean indicador = true;
    	try {
    		for (ParametroReporte parametroReporte : listaParametrosReporte) {
    			if (parametroReporte.getValorInicial()==null && parametroReporte.getValorInicial().equals("")
						  && parametroReporte.getValorFinal()==null && parametroReporte.getValorFinal().equals("") 
						  && CoreUtil.validarNumero(parametroReporte.getValorFinal())  && CoreUtil.validarNumero(parametroReporte.getValorFinal())){
    				indicador = false;
    				break;
    			}
			}
		} catch (Exception excepcion) {
			throw new GmdException(excepcion);
		}
		return indicador;
	}
	private List<List<String>> obtenerListaNumerosReclamo(ParametroReporte parametroReporte) throws GmdException {
		  List<List<String>> listaNumeroReclamos = new ArrayList<List<String>>();
		  try {
			  List<String> listaNumeroReclamosTemp = new ArrayList<String>();
			  Integer numeroReclamoInicial = Integer.valueOf(parametroReporte.getValorInicial());
			  Integer numeroReclamoFinal = Integer.valueOf(parametroReporte.getValorFinal());
			  Integer rangoNumeros = numeroReclamoFinal - numeroReclamoInicial +1;
			  
			  if (rangoNumeros <= cantNroReclamosConsulta) {
				  for (int numeroReclamo = numeroReclamoInicial ; numeroReclamo <= numeroReclamoFinal; numeroReclamo++) {
					  if (Constantes.ID_PARAM_FORMATO_NUMERO_RECLAMO_TIPO_NUMERICO.compareTo(parametroReporte.getIdFormtNroRec())==0) {
						  listaNumeroReclamosTemp.add(StringUtils.leftPad(String.valueOf(numeroReclamo), Constantes.CANTIDAD_CARACTERES_NUMERO_RECLAMO, "0"));
					  } else {
						  listaNumeroReclamosTemp.add(parametroReporte.getValorUnoGenerico() + StringUtils.leftPad(String.valueOf(numeroReclamo), Constantes.CANTIDAD_CARACTERES_NUMERO_RECLAMO, "0"));
					  }
				  }
				  listaNumeroReclamos.add(listaNumeroReclamosTemp);
			  } else {
				  Integer numeroInicial = numeroReclamoInicial;
				  for (int numeroReclamo = numeroReclamoInicial ; numeroReclamo <= numeroReclamoFinal; numeroReclamo++) {
					  if (Constantes.ID_PARAM_FORMATO_NUMERO_RECLAMO_TIPO_NUMERICO.compareTo(parametroReporte.getIdFormtNroRec())==0) {
						  listaNumeroReclamosTemp.add(StringUtils.leftPad(String.valueOf(numeroReclamo), Constantes.CANTIDAD_CARACTERES_NUMERO_RECLAMO, "0"));
					  } else {
						  listaNumeroReclamosTemp.add(parametroReporte.getValorUnoGenerico() + StringUtils.leftPad(String.valueOf(numeroReclamo), Constantes.CANTIDAD_CARACTERES_NUMERO_RECLAMO, "0"));
					  }
					  if (numeroReclamo == (numeroInicial +  cantNroReclamosConsulta -1) ) {
					  
						  listaNumeroReclamos.add(listaNumeroReclamosTemp);
						  listaNumeroReclamosTemp = new ArrayList<String>();
						  numeroInicial = numeroReclamo + 1;
					  } else {
						if (numeroReclamo == numeroReclamoFinal) {
							listaNumeroReclamos.add(listaNumeroReclamosTemp);
						}
					  }
				  }
			  }
			} catch (Exception excepcion) {
				throw new GmdException(excepcion);
			}
		  return listaNumeroReclamos;
	}
	public void generarReporteExcelSeguimientoResultado() {
		  try {
			  /*REQ001 - Item001 - Inicio */
			  frenteAtencionBean = new FrenteAtencion();
			  frenteBean = new FrenteAtencion();
			  /*REQ001 - Item001 - Fin */
			  int idFrente=0, idAtencion=0;
			  String frenteDesc="", atencionDesc="";
			  String cantidad = PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_ALFRESCO, "cantidad_numeros_reclamo_consulta");
			  cantNroReclamosConsulta = Integer.valueOf(cantidad);
			  Empleado usuarioSeguridad = (Empleado) WebUtil.obtenerObjetoSesion(Constantes.SESION_USUARIOINICIO);
			  FechaUtil fechaMes = new FechaUtil();
			  List<ParametroReporte> listaParaMetrosInterv = new ArrayList<ParametroReporte>();
			  Session sessionCmis = (Session)WebUtil.obtenerObjetoSesion(Constantes.SESION_SESIONCMIS);
			  String where, where02;
			  List<Map<String, Object>> listaMetadataReporte = new ArrayList<Map<String,Object>>();
			  List<Map<String, Object>> listaMetadataReporteNoExistentes = new ArrayList<Map<String,Object>>();
			  List<Map<String, Object>> listaMeses = new ArrayList<Map<String,Object>>();
			  List<Map<String, Object>> listaMetadataReporteErrores = new ArrayList<Map<String,Object>>();
			  List<Map<String, Object>> listaMetadataRepErrores = new ArrayList<Map<String,Object>>();
			  List<Map<String, Object>> listaMetadataReporteSeguimientoNoExistentes = new ArrayList<Map<String,Object>>();
			  
			  Map<String,Object> mapReporteSeguimiento = new HashMap<String,Object>();
			  List<Map<String, Object>> listaMetadataReporteSeguimientoTmp = new ArrayList<Map<String,Object>>();
			  List<String> listaGeneralNumeroSeguimiento = new ArrayList<String>();
			  List<List<String>> listaNumeroSeguimiento;
			  List<Map<String, Object>> listaReclamosSeguimientoExisten = new ArrayList<Map<String,Object>>();
			  
			  mapReporteSeguimiento.put("mes", seccionMes);
			  mapReporteSeguimiento.put("anio", seccionAnio);
			  
			  Map<String,Object> parametrosFrenteEquipo = new HashMap<String,Object>();
			  idFrente=reclamo.getIdFrenteAtencion();
			  idAtencion=reclamo.getIdEquipoAtencion();
			  frenteAtencionBean.setId(idFrente);
			  frenteAtencionBean =frenteAtencionService.listarFrenteAtencion(frenteAtencionBean);
			  frenteBean.setId(idAtencion);
			  equipoTrabajoBean=frenteAtencionService.equipoTrabajoListar(frenteBean);
			  atencionDesc=reclamo.getDescripcionEquipoAtencion();
			  
			  if (idFrente!=0){
				  parametrosFrenteEquipo.put("frente", idFrente);  
			  }else{
				  parametrosFrenteEquipo.put("frente", "0");
			  }
			  if (idAtencion!=0){
				  parametrosFrenteEquipo.put("equipo",idAtencion );
			  }else{
				  parametrosFrenteEquipo.put("equipo","0");
			  }
			  String inicioMesInterv=  seccionAnio + "-"+ seccionMes + "-01 00:00:00";
			  String finMesInterv=  seccionAnio + "-"+ seccionMes + "-31 00:00:00";
			  parametrosFrenteEquipo.put("inicioMesInterv", inicioMesInterv);
			  parametrosFrenteEquipo.put("finMesInterv", finMesInterv);
			  
			  //listaParaMetrosInterv trae el parametros de la opcion Parametros Reporte
			  if (!seccionMes.equals("")&& !seccionAnio.equals("")){
				  listaParaMetrosInterv = parametroReporteService.listarParametroReporteSeguimiento(mapReporteSeguimiento);
				  if (listaParaMetrosInterv!=null && !listaParaMetrosInterv.isEmpty()){
					  List<String> listaParametros = new ArrayList<String>();
					  List<String> listaParametrosAux = new ArrayList<String>();
					  if (validarParametroInicialFinal(listaParaMetrosInterv)) {
						  listaGeneralNumeroSeguimiento = generarListaReporte(listaParaMetrosInterv);
						  for(ParametroReporte valoresReporte: listaParaMetrosInterv ){
							  listaNumeroSeguimiento = new ArrayList<List<String>>();
							  listaNumeroSeguimiento = obtenerListaNumerosReclamo(valoresReporte);
							  for (List<String> numeroReclamosConsulta : listaNumeroSeguimiento) {
								  where = WebUtil.obtenerReporteWhereCmis(numeroReclamosConsulta);
								  listaMetadataReporteSeguimientoTmp = alfrescoService.listarObjetosFiltroReporte(sessionCmis, TablasQueryEnumSisRec.BBVAFOLDER.toString(), where, parametrosFrenteEquipo);
								  if (listaMetadataReporteSeguimientoTmp != null && !listaMetadataReporteSeguimientoTmp.isEmpty()) {
									listaReclamosSeguimientoExisten.addAll(listaMetadataReporteSeguimientoTmp);
								  }
							  }
						  }
					  }else {
							WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_ValoresInicialFinal"));
							return;
					  }
					  
					  // Recuperar files no existentes
					  if (idFrente==0){
						  listaMetadataReporteSeguimientoNoExistentes = obtenerListaReclamosNoExisten(listaGeneralNumeroSeguimiento,listaReclamosSeguimientoExisten);  
					  }
					   
					  // Recuperar files Errores
					  Map<String, Object> properties = new HashMap<String, Object>() ;
					  String inicioMes=  seccionAnio + "-"+ seccionMes + "-01 00:00:00";
					  String finMes=  seccionAnio + "-"+ seccionMes + "-31 00:00:00";
					  properties.put("inicioMes", inicioMes);
					  properties.put("finMes", finMes);
					  listaMeses.add(properties);
					  where = inicioMes;
					  where02=finMes;
					  Map<String,Object> parametrosFrenteEquipoError = new HashMap<String,Object>();
					  if (idFrente!=0){
				  		parametrosFrenteEquipoError.put("frente", idFrente);  
					  }else{
						  parametrosFrenteEquipoError.put("frente", "0");
					  }
					  if (idAtencion!=0){
						  parametrosFrenteEquipoError.put("equipo",idAtencion );
					  }else{
						  parametrosFrenteEquipoError.put("equipo","0");
					  }
					  listaMetadataReporteErrores= alfrescoService.listarObjetosReporte(sessionCmis, TablasQueryEnumSisRec.BBVAFOLDER.toString(), where, where02, parametrosFrenteEquipoError);
					  Map<String, Object> prop = new HashMap<String, Object>() ;
					  for (int i = 0; i < listaMetadataReporteErrores.size(); i++) {
						  boolean encuentra = false;
						  for (int j = 0; j <listaReclamosSeguimientoExisten.size(); j++) {
							  prop = new HashMap<>();
							  if (listaMetadataReporteErrores.get(i).get("Nro Reclamo").toString().equals(listaReclamosSeguimientoExisten.get(j).get("Nro Reclamo").toString())) {
								  encuentra = true;
								  break;
							  }
						  }
						  if(!encuentra){
							  listaMetadataRepErrores.add(listaMetadataReporteErrores.get(i));
						  }
				  	 }
				 
					// Reporte Excel
					LOGGER.info("Inicio generarReporteExcel");
					HttpServletResponse response = (HttpServletResponse) FacesContext.getCurrentInstance().getExternalContext().getResponse();
					InputStream flujoBytesExcel=null;
				
					flujoBytesExcel = getClass().getResourceAsStream("/plantillaReporteExcel.xls");
					HSSFWorkbook libroExcel = new HSSFWorkbook(flujoBytesExcel);
					HSSFSheet hojaExcel = libroExcel.getSheetAt(0);
				  
					UDocumentoExcel uDocumentoExcel = new UDocumentoExcel(libroExcel);
					HSSFCellStyle csCeldaFilaBlanca = uDocumentoExcel.generarEstiloCeldaTablaBlanco(HSSFCellStyle.ALIGN_LEFT);
					HSSFCellStyle csCeldaFilaColor = uDocumentoExcel.generarEstiloCeldaTablaColor(HSSFCellStyle.ALIGN_LEFT);
				  
					HSSFRow filaExcel = null;
					HSSFCell celdaExcel = null;
				  
					
					  Integer totalRegistros  = (listaReclamosSeguimientoExisten.size() + listaMetadataReporteNoExistentes.size() + listaMetadataRepErrores.size());

					
					/* Ancho por defecto de las columnas; importante no se configura el ancho de la celda 0 ya que esta viene de la plantilla excel por defecto. */
					hojaExcel.setColumnWidth((short)(1),(short)4000);
					hojaExcel.setColumnWidth((short)(2),(short)5000); 
					hojaExcel.setColumnWidth((short)(3),(short)4000);
					hojaExcel.setColumnWidth((short)(4),(short)6000);
					hojaExcel.setColumnWidth((short)(5),(short)10000);
		         
					filaExcel = hojaExcel.createRow(4);
					celdaExcel = filaExcel.createCell((short)0);

					//Insertar la fecha
					filaExcel = hojaExcel.getRow(1);
					celdaExcel = filaExcel.getCell((short)7);
					Calendar fechaActual = Calendar.getInstance();
					String strFechaActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.DAY_OF_MONTH)+"",2)+"/"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MONTH)+1)+"",2)+"/"+Funciones.CompletaCerosIzq(fechaActual.get(Calendar.YEAR)+"",4);
					celdaExcel.setCellValue(strFechaActual);

					// Insertar la Hora
					filaExcel = hojaExcel.getRow(2);
					celdaExcel = filaExcel.getCell((short)7);
					String sHoraActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.HOUR)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MINUTE)+1)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.SECOND)+1)+"",2);
					celdaExcel.setCellValue(sHoraActual);
		          
					// Insertar el Usuario
					filaExcel = hojaExcel.getRow(3);
					celdaExcel = filaExcel.getCell((short)7);
					celdaExcel.setCellValue( usuarioSeguridad.getLogin());
					
					int  mesActual=0 , anioActual=0;
					String fechaTempCabecera[]=  strFechaActual.split("/");
					mesActual=Integer.valueOf(fechaTempCabecera[1]);
					anioActual=Integer.valueOf(fechaTempCabecera[2]);
					
					filaExcel = hojaExcel.createRow(5);
					celdaExcel = filaExcel.createCell((short)0);
					celdaExcel.setCellValue("Total Registros encontrados = " + (listaReclamosSeguimientoExisten.size() + listaMetadataReporteSeguimientoNoExistentes.size() + listaMetadataRepErrores.size()));
					LOGGER.info("listaDocumentos :: "+listaMetadataReporte.size());
					String nombreReporteAnioMes="";
					nombreReporteAnioMes= "Periodo: " +(fechaUtil.obtenerNombreMes(Integer.valueOf(mesActual) ) )+ " - " + anioActual;
					if (idFrente!=0){
						nombreReporteAnioMes= nombreReporteAnioMes + "  - "+ frenteAtencionBean.getDescripcion();
						
					}
					if (idAtencion!=0){
						nombreReporteAnioMes= nombreReporteAnioMes + " -  "+ equipoTrabajoBean.getDescripcion();
					}
					celdaExcel = filaExcel.createCell((short)3);
					celdaExcel.setCellValue(nombreReporteAnioMes);
					
			       //Insertar en nombre del reporte
						filaExcel = hojaExcel.getRow(2);
						celdaExcel = filaExcel.getCell((short)3);
						celdaExcel.setCellValue("REPORTE - SEGUIMIENTO RESULTADO");
		          
					/* Fila desde la que se inicia. */
					short numeroFila = 7;
					int contaReporte=1;
		          
					// Datos de la tabla detalle:
					//LISTA DE EXISTE - listaReclamosSeguimientoExisten
					 short numeroFilaTotal = 0;
					if(listaReclamosSeguimientoExisten!=null && !listaReclamosSeguimientoExisten.isEmpty()){
						String strTemporal = "";  
						int contador=0;
						int  mes=0;
						String fecha="", codUsuarioRegistro,  nombreCompleto="";
						Map<String,Object> parametersUsuario = new HashMap<String,Object>();   
						UsuarioReporte usuario = new UsuarioReporte();
						for (Map<String, Object>  reporteSeguimiento : listaReclamosSeguimientoExisten ) {
							
							numeroFilaTotal ++;
							if(numeroFilaTotal > Constantes.NUMERO_MAXIMO_FILAS_EXCEL){
								numeroFilaTotal = 0;
								hojaExcel=crearHojaExcelReporteCierreMes(libroExcel, "Reporte Seguimiento Resultado", nombreReporteAnioMes, totalRegistros,"REPORTE - SEGUIMIENTO RESULTADO");
								numeroFila = 7;
								}
							
							
							//Inicio obtener Mes 
							if (!reporteSeguimiento.get("Fecha Registro Reclamo").toString().equals("") && reporteSeguimiento.get("Fecha Registro Reclamo").toString()!=null){
								fecha= reporteSeguimiento.get("Fecha Registro Reclamo").toString();
								String fechaTemp[]=  fecha.split("/");
								mes=Integer.valueOf(fechaTemp[1]);
							}
							//Fin obtener Mes 
							
							//Inicio traer nombre Usuario
							if (!reporteSeguimiento.get("Cod Registro").toString().equals("") && reporteSeguimiento.get("Cod Registro").toString()!=null){
								usuario.setLogin(reporteSeguimiento.get("Cod Registro").toString());
								usuario =usuarioService.nombreUsuario(usuario);
								if (usuario!=null && usuario.getNombre()!=null && !usuario.getNombre().equals("")){
									usuario.setNombreCompleto(usuario.getNombre() + " " + usuario.getApellidoPaterno() + " " + usuario.getApellidoMaterno() );
								}else{
									usuario=new UsuarioReporte();
									usuario.setNombreCompleto(" - ");
								}
							}
							//Fin traer nombre Usuario
		          		
							/* Estilo de la fila. */
							HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
		          		
			          		filaExcel = hojaExcel.createRow(++numeroFila);
			                celdaExcel = filaExcel.createCell((short)0);
			                strTemporal = String.valueOf( contaReporte );
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			          		
			                celdaExcel = filaExcel.createCell((short)1);
			                strTemporal =  reporteSeguimiento.get("Nro Reclamo").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)2);
			                strTemporal =  reporteSeguimiento.get("Estado").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)3);
			                strTemporal = reporteSeguimiento.get("Frente").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)4);
			                strTemporal = reporteSeguimiento.get("Equipo").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)5); 
			                strTemporal =  usuario.getNombreCompleto();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)6);
			                strTemporal =  fechaMes.obtenerNombreMes(mes);
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)7);
			                strTemporal =  reporteSeguimiento.get("Fecha Registro Reclamo").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                /*REQ001 - Item001 - Inicio*/
			                celdaExcel = filaExcel.createCell((short)8);
			                strTemporal =  reporteSeguimiento.get("Fecha Modificación").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                /*REQ001 - Item001 - Fin*/
			                contaReporte++;
						}
					}
					// NO EXISTE - listaMetadataReporteSeguimientoNoExistentes
					if (idFrente==0){
						if(listaMetadataReporteSeguimientoNoExistentes!=null && !listaMetadataReporteSeguimientoNoExistentes.isEmpty()){
							for (Map<String, Object>  reporteNoExistentes : listaMetadataReporteSeguimientoNoExistentes ) {
								numeroFilaTotal ++;
								if(numeroFilaTotal > Constantes.NUMERO_MAXIMO_FILAS_EXCEL){
									numeroFilaTotal = 0;
									hojaExcel=crearHojaExcelReporteCierreMes(libroExcel, "Reporte Seguimiento Resultado ", nombreReporteAnioMes, totalRegistros,"REPORTE - SEGUIMIENTO RESULTADO");
									numeroFila = 7;
									}
								
								
								String strTemporal = "";  
			        			int contador=0;
			        			/* Estilo de la fila. */
				          		HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
				          		
				          		filaExcel = hojaExcel.createRow(++numeroFila);
				                celdaExcel = filaExcel.createCell((short)0);
				                strTemporal = String.valueOf( contaReporte );
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				          		
				                celdaExcel = filaExcel.createCell((short)1);
				                strTemporal =  reporteNoExistentes.get("Nro Reclamo").toString();
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                
				                celdaExcel = filaExcel.createCell((short)2);
				                strTemporal =  "NO EXISTE";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                
				                celdaExcel = filaExcel.createCell((short)3);
				                strTemporal = " - ";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				               
				                celdaExcel = filaExcel.createCell((short)4); 
				                strTemporal =  " - ";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				               
				                celdaExcel = filaExcel.createCell((short)5);
				                strTemporal =   " - ";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                
				                celdaExcel = filaExcel.createCell((short)6);
				                strTemporal =   " - ";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                
				                celdaExcel = filaExcel.createCell((short)7);
				                strTemporal =   " - ";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                
				                /*REQ001 - Item001 - Inicio*/
				                celdaExcel = filaExcel.createCell((short)8);
				                strTemporal =   " - ";
				                celdaExcel.setCellValue(strTemporal);
				                celdaExcel.setCellStyle(estilo);
				                /*REQ001 - Item001 - Fin*/
				                
				                contaReporte++;
							}
						}
					}
		          
					//ERROR - listaMetadataRepErrores
					if(listaMetadataRepErrores.size()>0 && listaMetadataRepErrores!=null){
		        		UsuarioReporte usuario = new UsuarioReporte();
		        		String fecha="";
		        		int  mes=0;
		        		for (Map<String, Object>  reporteErrores : listaMetadataRepErrores ) {
		        			numeroFilaTotal ++;
							if(numeroFilaTotal > Constantes.NUMERO_MAXIMO_FILAS_EXCEL){
								numeroFilaTotal = 0;
								hojaExcel=crearHojaExcelReporteCierreMes(libroExcel, "Reporte Seguimiento Resultado ", nombreReporteAnioMes, totalRegistros,"REPORTE - SEGUIMIENTO RESULTADO");
								numeroFila = 7;
								}
							
		        			//Inicio obtener Mes 
			          		if (!reporteErrores.get("Fecha Registro Reclamo").toString().equals("") && reporteErrores.get("Fecha Registro Reclamo").toString()!=null){
			          			fecha= reporteErrores.get("Fecha Registro Reclamo").toString();
				          		String fechaTemp[]=  fecha.split("/");
				          		mes=Integer.valueOf(fechaTemp[1]);
			          		}
			          		//Fin obtener Mes 
			          		
			          		if (!reporteErrores.get("Cod Registro").toString().equals("") && reporteErrores.get("Cod Registro").toString()!=null){
			          			usuario.setLogin(reporteErrores.get("Cod Registro").toString());
			          			usuario =usuarioService.nombreUsuario(usuario);
			          			if (usuario!=null && usuario.getNombre()!=null && !usuario.getNombre().equals("")){
			          				usuario.setNombreCompleto(usuario.getNombre() + " " + usuario.getApellidoPaterno() + " " + usuario.getApellidoMaterno() );
			          			}else{
			          				usuario=new UsuarioReporte();
			          				usuario.setNombreCompleto(" - ");
			          			}
			          		}
		        		  
			          		String strTemporal = "";
		        			int contador=0;
		        			/* Estilo de la fila. */
			          		HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
			          		
			          		filaExcel = hojaExcel.createRow(++numeroFila);
			                celdaExcel = filaExcel.createCell((short)0);
			                strTemporal = String.valueOf( contaReporte );
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			          		
			                celdaExcel = filaExcel.createCell((short)1);
			                strTemporal =  reporteErrores.get("Nro Reclamo").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)2);
			                strTemporal = "ERROR - " +  reporteErrores.get("Estado").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)3);
			                strTemporal = reporteErrores.get("Frente").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)4);
			                strTemporal = reporteErrores.get("Equipo").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)5); 
			                strTemporal =  usuario.getNombreCompleto();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)6);
			                strTemporal =  fechaMes.obtenerNombreMes(mes);
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                celdaExcel = filaExcel.createCell((short)7);
			                strTemporal =  reporteErrores.get("Fecha Registro Reclamo").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                
			                /*REQ001 - Item001 - Inicio*/
			                celdaExcel = filaExcel.createCell((short)8);
			                strTemporal =  reporteErrores.get("Fecha Modificación").toString();
			                celdaExcel.setCellValue(strTemporal);
			                celdaExcel.setCellStyle(estilo);
			                /*REQ001 - Item001 - Fin*/
			                contaReporte++;
		        		}
					}
		          
					/* Pie de pagina del reporte. */
					numeroFila=(short)(numeroFila + 2);
					filaExcel = hojaExcel.createRow(numeroFila);
					celdaExcel = filaExcel.createCell((short)0);
					celdaExcel.setCellStyle(uDocumentoExcel.generarEstiloLeyenda(HSSFCellStyle.ALIGN_CENTER, false));
					celdaExcel.setCellValue("****FIN DE REPORTE****");
					  hojaExcel.addMergedRegion(new Region(numeroFila, (short) 0, numeroFila, (short) 1));
					  
					/* Enviar archivo al navegador para su descarga. */
					response.setContentType("application/vnd.ms-excel");
					response.setHeader("Content-Disposition","attachment; filename=\"ReporteSeguimientoResultado.xls\"");
					OutputStream out = response.getOutputStream();
					libroExcel.write(out);
					out.flush();
					out.close();
					FacesContext.getCurrentInstance().responseComplete();
				  }else{
					  WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_listaVacia"));
				  }
				}else{
				  WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_seleccionAnioMes"));
				}
	    	}catch(Exception exception){
	    		LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(exception)[1], exception);
	 	    	String[] error = MensajeExceptionUtil.obtenerMensajeError(exception);
	 	    	String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
	 	    	WebUtil.mensajeError(mensajeError);
	    	}
		  LOGGER.info("Fin generarReporteExcel");
	}
	   
	  /**
	   * Es la lista que ira para el WHERE del CMIS
	   */
	  public List<String> generarListaReporte(List<ParametroReporte> parametrosReporte){
		  String numerosReclamo = "";
		  int numero=0;
		  List<String> valor = new ArrayList<String>();
		  for (ParametroReporte parametroReporte : parametrosReporte) {
			  if (parametroReporte.getValorInicial()!=null && !parametroReporte.getValorInicial().equals("")
				  && parametroReporte.getValorFinal()!=null && !parametroReporte.getValorFinal().equals("") ){
				  if (Constantes.ID_PARAM_FORMATO_NUMERO_RECLAMO_TIPO_NUMERICO.compareTo(parametroReporte.getIdFormtNroRec())==0) {
					  String valorInicial=parametroReporte.getValorInicial();
					  String valorFinal=parametroReporte.getValorFinal();
					  for (int numeroReclamo = Integer.parseInt(valorInicial); numeroReclamo <= Integer.parseInt(valorFinal); numeroReclamo++) {
						numerosReclamo=String.valueOf(numeroReclamo);
						valor.add(StringUtils.leftPad(numerosReclamo, Constantes.CANTIDAD_CARACTERES_NUMERO_RECLAMO, "0"));
					}
				  } else {
					String valorInicial=parametroReporte.getValorInicial();
					String valorFinal=parametroReporte.getValorFinal();
					for (int numeroReclamo = Integer.parseInt(valorInicial); numeroReclamo <= Integer.parseInt(valorFinal); numeroReclamo++) {
						numerosReclamo=String.valueOf(numeroReclamo);
						valor.add(parametroReporte.getValorUnoGenerico() + StringUtils.leftPad(numerosReclamo, Constantes.CANTIDAD_CARACTERES_NUMERO_RECLAMO, "0"));
					}
				  }
			  }
		  }
		  return valor;
	  }
	  
	  public String irReporteEstadoActual(){
			String ruta = "";
			try {
				ruta = "reporteEstadoActual?faces-redirect=true";
			} catch (Exception exception) {
				LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(exception)[1], exception);
		    	String[] error = MensajeExceptionUtil.obtenerMensajeError(exception);
		    	String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
		    	WebUtil.mensajeError(mensajeError);
			}
			return ruta;
	  }
	  
	  public void generarReporteEstadoActual() throws GmdException {
		  try {
				
			  /*REQ001 - Item001 - Inicio */
			  
			if(reclamo.getFechaDesde()!=null && reclamo.getFechaHasta()!=null ){
				  			  
				
				
			  Calendar fechaInicioC = Calendar.getInstance();
			  Calendar fechaFinC = Calendar.getInstance();
			  fechaInicioC.setTime(reclamo.getFechaDesde());
			  fechaFinC.setTime(reclamo.getFechaHasta());
						  
			  int feIni=fechaInicioC.get(Calendar.MONTH)+1;
						     
			  int feFin=fechaFinC.get(Calendar.MONTH)+1;
				  
			  if(  feIni - feFin <= 0){
					 
				  /*REQ001 - Item001 - Fin */ 
			
			  //PR
			  Empleado usuarioSeguridad = (Empleado) WebUtil.obtenerObjetoSesion(Constantes.SESION_USUARIOINICIO);
			  String anioActualReporte;
			  int anioPasado;
			  Date fechaActualReporte = new Date();
			  String dateStringReporte = fechaUtil.obtenerFechaStr(fechaActualReporte);
			  String anioTempReporte[]= dateStringReporte.split("/");
			  anioActualReporte=anioTempReporte[2];
			  anioPasado=Integer.parseInt(anioActualReporte)- 1;
			
			  List<Reclamo> listaReclamosHistorico = new ArrayList<Reclamo>();
			  List<Reclamo> listaReclamosMes = new ArrayList<Reclamo>();
			  Map<String, Object> listaMatriz = new HashMap<String, Object>();
			  int totalAnioPasado = 0, totalAnioActualCerrado=0, totalAnioActualAbierto=0, totalAnioActualNoExiste=0, totalAnioActualError=0;
			  Map<String, Object> listaAnioActualMes = new HashMap<String,Object>();
			  Reclamo reclamo= new Reclamo();
			  reclamo.setAnio(String.valueOf(anioPasado));
			  listaReclamosHistorico = reporteService.listarDatosAnioPasado(reclamo);
			  ///INICIO
			  List<Reclamo> listaReporteAnioPasado = obtenerListaReclamosHistorico(listaReclamosHistorico);
			  ///FIN
			  
			  //Inicio de obtener la lista de los meses del anio actual
			  int contaMes=0;
			  String descMes;List<String> listaMes = new ArrayList<String>();
			  
	    
			  reclamo.setAnio(String.valueOf(anioActualReporte));
			  /*REQ001 - Item001 - Inicio */
			  for (int i = feIni; i <= feFin; i++) {   	 
				     
				  if(contaMes <=9){
					  reclamo.setMes("0"+ i);
					  descMes=fechaUtil.obtenerNombreMes(Integer.parseInt("0"+ i));
				  }else{
					  reclamo.setMes(""+ i);
					  descMes=fechaUtil.obtenerNombreMes(Integer.parseInt(""+ i));
				  }
				  
				  listaReclamosMes=reporteService.listarDatosAnioActual(reclamo);
				  if (listaReclamosMes!=null && listaReclamosMes.size()>0){
					  //INICIO
					  List<Reclamo> listaAnioActual = obtenerListaReclamosHistorico(listaReclamosMes);
					  //FIN
					  listaMatriz.put(descMes,listaAnioActual );
					  listaMes.add(descMes);
				  }
				  contaMes++;
		     
		     
		     }
			  /*REQ001 - Item001 - Fin */
			  ///
			  /* Original
			  for (int i = 1; i <= 12; i++) {
				  if(contaMes <=9){
					  reclamo.setMes("0"+ i);
					  descMes=fechaUtil.obtenerNombreMes(Integer.parseInt("0"+ i));
				  }else{
					  reclamo.setMes(""+ i);
					  descMes=fechaUtil.obtenerNombreMes(Integer.parseInt(""+ i));
				  }
				  listaReclamosMes=reporteService.listarDatosAnioActual(reclamo);
				  if (listaReclamosMes!=null && listaReclamosMes.size()>0){
					  //INICIO
					  List<Reclamo> listaAnioActual = obtenerListaReclamosHistorico(listaReclamosMes);
					  //FIN
					  listaMatriz.put(descMes,listaAnioActual );
					  listaMes.add(descMes);
				  }
				  contaMes++;
			  }
			  */
			  /******/
			  int cantTotalMes01=0, cantTotalMes02=0,cantTotalMes03=0,cantTotalMes04=0,cantTotalMes05=0;
			  int cantTotalAnio01=0,cantTotalAnio02=0,cantTotalAnio03=0,cantTotalAnio04=0,cantTotalAnio05=0;
			  
			  //Fin de obtener la lista de los meses del anio actual
			  
			  // Reporte Excel
			  HttpServletResponse response = (HttpServletResponse) FacesContext.getCurrentInstance().getExternalContext().getResponse();
			  InputStream flujoBytesExcel=null;
			  flujoBytesExcel = getClass().getResourceAsStream("/plantillaReporte.xls");
			  HSSFWorkbook libroExcel = new HSSFWorkbook(flujoBytesExcel);
			  HSSFSheet hojaExcel = libroExcel.getSheetAt(0);
			  
			  UDocumentoExcel uDocumentoExcel = new UDocumentoExcel(libroExcel);
			  HSSFCellStyle csCeldaFilaBlanca = uDocumentoExcel.generarEstiloCeldaTablaBlanco(HSSFCellStyle.ALIGN_LEFT);
			  HSSFCellStyle csCeldaFilaColor = uDocumentoExcel.generarEstiloCeldaTablaColor(HSSFCellStyle.ALIGN_LEFT);
			  HSSFCellStyle estiloCabeceraTabla = uDocumentoExcel.generarEstiloCabeceraTabla();
			  
			  HSSFRow filaExcel = null;
			  HSSFCell celdaExcel = null;
			  
			  /* Ancho por defecto de las columnas; importante no se configura el ancho de la celda 0 ya que esta viene de la plantilla excel por defecto. */
			  hojaExcel.setColumnWidth((short)(1),(short)4000);
			  hojaExcel.setColumnWidth((short)(2),(short)4000); 
			  hojaExcel.setColumnWidth((short)(3),(short)4000);
			  hojaExcel.setColumnWidth((short)(4),(short)8000);
			  
			  filaExcel = hojaExcel.createRow(4);
			  celdaExcel = filaExcel.createCell((short)0);
			  		 
			  //Insertar en nombre del reporte
			  filaExcel = hojaExcel.getRow(2);
			  celdaExcel = filaExcel.getCell((short)2);
			  celdaExcel.setCellValue("REPORTE - ESTADO ACTUAL");
			  //Insertar la fecha
			  filaExcel = hojaExcel.getRow(1);
			  celdaExcel = filaExcel.getCell((short)6);
			  Calendar fechaActual = Calendar.getInstance();
			  String strFechaActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.DAY_OF_MONTH)+"",2)+"/"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MONTH)+1)+"",2)+"/"+Funciones.CompletaCerosIzq(fechaActual.get(Calendar.YEAR)+"",4);
			  celdaExcel.setCellValue(strFechaActual);
			  
			  // Insertar la Hora
			  filaExcel = hojaExcel.getRow(2);
			  celdaExcel = filaExcel.getCell((short)6);
			  String sHoraActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.HOUR)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MINUTE)+1)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.SECOND)+1)+"",2);
			  celdaExcel.setCellValue(sHoraActual);
			  
			  // Insertar el Usuario
			  filaExcel = hojaExcel.getRow(3);
			  celdaExcel = filaExcel.getCell((short)6);
			  celdaExcel.setCellValue( usuarioSeguridad.getLogin());
			  
			  
			  filaExcel = hojaExcel.getRow(6);
			  celdaExcel = filaExcel.getCell((short)1);
			  String reporteAnioPasado = "";
			  celdaExcel.setCellValue(reporteAnioPasado);
			      
			   // Insertamos el anio pasado
			  filaExcel = hojaExcel.getRow(7);
			  celdaExcel = filaExcel.getCell((short)1);
			  String anioPasadoReporte = "TOTAL "+ String.valueOf(anioPasado);
			  celdaExcel.setCellValue(anioPasadoReporte);
			  celdaExcel.setCellStyle(estiloCabeceraTabla);
			  
			  // Insertamos el anio actual
			  int numTemp = 2;
			  for (String string : listaMes) {
				  filaExcel = hojaExcel.getRow(7);
				  celdaExcel = filaExcel.getCell((short)numTemp);
				  celdaExcel.setCellValue(string + " " + String.valueOf(anioActualReporte));
				  celdaExcel.setCellStyle(estiloCabeceraTabla);
				  numTemp ++;
			  }
			  // Insertamos el texto: TOTAL + anio actual
			  filaExcel = hojaExcel.getRow(7);
			  celdaExcel = filaExcel.getCell((short)numTemp);
			  String TotalGeneral = "TOTAL "  + String.valueOf(anioActualReporte);
			  celdaExcel.setCellValue(TotalGeneral);
			  celdaExcel.setCellStyle(estiloCabeceraTabla);
			  /* Fila desde la que se inicia. */
			
			  int contaReporte=1,totalReporte01=0,totalReporte02=0,totalReporte03=0,totalReporte04=0,totalReporte05=0,totalReporte06=0,totalReporte07=0,totalReporte08=0,totalReporte09=0,totalReporte10=0,totalReporte11=0,totalReporte12=0 ;
			  short numeroFila = 7;
			  
			  // Datos de la tabla detalle:
			  if(listaMatriz!=null && listaMatriz.size()>0){
				  /***INICIO DETALLE - EXCEL**/
				  int contador1=0;
				  int i=0;
				  for (Reclamo objReporteAnioPasadoj :  listaReporteAnioPasado){
					  String strTemporal = "";  
					  HSSFCellStyle estilo = ((contador1%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
				  
					  filaExcel = hojaExcel.createRow(++numeroFila);
					  int num = 2;
					  if (i == 0) { // FILE ABIERTO
						  celdaExcel = filaExcel.createCell((short)0);
					      celdaExcel.setCellValue("FILE ABIERTO");
					      celdaExcel.setCellStyle(estilo);
					      
						  cantTotalAnio02 = obtenerCantidadReclamosAnio(listaReporteAnioPasado, Constantes.ID_ESTADO_RECLAMO_FILE_ABIERTO);
						  celdaExcel = filaExcel.createCell((short)1);
					      totalAnioPasado+=objReporteAnioPasadoj.getTotales();
					      celdaExcel.setCellValue(cantTotalAnio02);
					      celdaExcel.setCellStyle(estilo);
						  for (String string : listaMes) {
							  cantTotalMes02 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_ABIERTO);
							  celdaExcel = filaExcel.createCell((short)num);
							  strTemporal = String.valueOf(cantTotalMes02);
							  totalAnioActualAbierto+=cantTotalMes02;
							  celdaExcel.setCellValue(strTemporal);
							  celdaExcel.setCellStyle(estilo);
							  num ++;
							  if(string.equals("ENERO")){
					        	totalReporte01=totalReporte01+cantTotalMes02;
							  }else if (string.equals("FEBRERO")){
					        	totalReporte02=totalReporte02+cantTotalMes02;
							  }else if (string.equals("MARZO")){
					        	totalReporte03=totalReporte03+cantTotalMes02;
							  }else if (string.equals("ABRIL")){
					        	totalReporte04=totalReporte04+cantTotalMes02;
							  }else if (string.equals("MAYO")){
					        	totalReporte05=totalReporte05+cantTotalMes02;
							  }else if (string.equals("JUNIO")){
					        	totalReporte06=totalReporte06+cantTotalMes02;
							  }else if (string.equals("JULIO")){
					        	totalReporte07=totalReporte07+cantTotalMes02;
							  }else if (string.equals("AGOSTO")){
					        	totalReporte08=totalReporte08+cantTotalMes02;
							  }else if (string.equals("SETIEMBRE")){
					        	totalReporte09=totalReporte09+cantTotalMes02;
							  }else if (string.equals("OCTUBRE")){
					        	totalReporte10=totalReporte10+cantTotalMes02;
							  }else if (string.equals("NOVIEMBRE")){
					        	totalReporte11=totalReporte11+cantTotalMes02;
							  }else if (string.equals("DICIEMBRE")){
					        	totalReporte12=totalReporte12+cantTotalMes02;
							  }
						  }
						  celdaExcel = filaExcel.createCell((short)num);
					      strTemporal = String.valueOf(totalAnioActualAbierto);
					      celdaExcel.setCellValue(strTemporal);
					      celdaExcel.setCellStyle(estilo);
				      //i++;
					  }else if (i == 1) { // FILE CERRADO
						  celdaExcel = filaExcel.createCell((short)0);
					      celdaExcel.setCellValue("FILE CERRADO");
					      celdaExcel.setCellStyle(estilo);
						  cantTotalAnio01 = obtenerCantidadReclamosAnio(listaReporteAnioPasado, Constantes.ID_ESTADO_RECLAMO_FILE_CERRADO);
						  celdaExcel = filaExcel.createCell((short)1);
					      totalAnioPasado+=objReporteAnioPasadoj.getTotales();
					      celdaExcel.setCellValue(cantTotalAnio01);
					      celdaExcel.setCellStyle(estilo);
				      
						  for (String string : listaMes) {
							  cantTotalMes01 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_CERRADO);
							  celdaExcel = filaExcel.createCell((short)num);
							  strTemporal = String.valueOf(cantTotalMes01);///27
							  totalAnioActualCerrado+=cantTotalMes01;
							  celdaExcel.setCellValue(strTemporal);
							  celdaExcel.setCellStyle(estilo);
							  num ++;
							  if(string.equals("ENERO")){
								  totalReporte01=totalReporte01+cantTotalMes01;
							  }else if (string.equals("FEBRERO")){
								  totalReporte02=totalReporte02+cantTotalMes01;
							  }else if (string.equals("MARZO")){
					        	totalReporte03=totalReporte03+cantTotalMes01;
							  }else if (string.equals("ABRIL")){
								  totalReporte04=totalReporte04+cantTotalMes01;
							  }else if (string.equals("MAYO")){
								  totalReporte05=totalReporte05+cantTotalMes01;
							  }else if (string.equals("JUNIO")){
								  totalReporte06=totalReporte06+cantTotalMes01;
							  }else if (string.equals("JULIO")){
								  totalReporte07=totalReporte07+cantTotalMes01;
							  }else if (string.equals("AGOSTO")){
								  totalReporte08=totalReporte08+cantTotalMes01;
							  }else if (string.equals("SETIEMBRE")){
								  totalReporte09=totalReporte09+cantTotalMes01;
							  }else if (string.equals("OCTUBRE")){
								  totalReporte10=totalReporte10+cantTotalMes01;
							  }else if (string.equals("NOVIEMBRE")){
								  totalReporte11=totalReporte11+cantTotalMes01;
							  }else if (string.equals("DICIEMBRE")){
								  totalReporte12=totalReporte12+cantTotalMes01;
							  }
						  }
						  celdaExcel = filaExcel.createCell((short)num);
					      strTemporal = String.valueOf(cantTotalMes02);
					      celdaExcel.setCellValue(totalAnioActualCerrado);
					      celdaExcel.setCellStyle(estilo);
					  }else if (i == 2) { // ERROR
						  celdaExcel = filaExcel.createCell((short)0);
					      celdaExcel.setCellValue("ERROR");
					      celdaExcel.setCellStyle(estilo);
					      
						  cantTotalAnio04 = obtenerCantidadReclamosAnio(listaReporteAnioPasado, Constantes.ID_ESTADO_RECLAMO_FILE_ERROR);
						  celdaExcel = filaExcel.createCell((short)1);
					      totalAnioPasado+=objReporteAnioPasadoj.getTotales();
					      celdaExcel.setCellValue(cantTotalAnio04);
					      celdaExcel.setCellStyle(estilo);
						  for (String string : listaMes) {
							  cantTotalMes04 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_ERROR);
							  celdaExcel = filaExcel.createCell((short)num);
					          strTemporal = String.valueOf(cantTotalMes04);
					          totalAnioActualError+=cantTotalMes04;
					          celdaExcel.setCellValue(strTemporal);
					          celdaExcel.setCellStyle(estilo);
					          num ++;
					          if(string.equals("ENERO")){
					        	totalReporte01=totalReporte01+cantTotalMes04;
					          }else if (string.equals("FEBRERO")){
					        	totalReporte02=totalReporte02+cantTotalMes04;
					          }else if (string.equals("MARZO")){
					        	totalReporte03=totalReporte03+cantTotalMes04;
					          }else if (string.equals("ABRIL")){
					        	totalReporte04=totalReporte04+cantTotalMes04;
					          }else if (string.equals("MAYO")){
					        	totalReporte05=totalReporte05+cantTotalMes04;
					          }else if (string.equals("JUNIO")){
					        	totalReporte06=totalReporte06+cantTotalMes04;
					          }else if (string.equals("JULIO")){
					        	totalReporte07=totalReporte07+cantTotalMes04;
					          }else if (string.equals("AGOSTO")){
					        	totalReporte08=totalReporte08+cantTotalMes04;
					          }else if (string.equals("SETIEMBRE")){
					        	totalReporte09=totalReporte09+cantTotalMes04;
					          }else if (string.equals("OCTUBRE")){
					        	totalReporte10=totalReporte10+cantTotalMes04;
					          }else if (string.equals("NOVIEMBRE")){
					        	totalReporte11=totalReporte11+cantTotalMes04;
					          }else if (string.equals("DICIEMBRE")){
					        	totalReporte12=totalReporte12+cantTotalMes04;
					          }
						  }
						  celdaExcel = filaExcel.createCell((short)num);
					      strTemporal = String.valueOf(totalAnioActualError);
					      celdaExcel.setCellValue(strTemporal);
					      celdaExcel.setCellStyle(estilo);
					  }else if (i == 3) {// NO EXISTE
						  celdaExcel = filaExcel.createCell((short)0);
					      celdaExcel.setCellValue("NO EXISTE");
					      celdaExcel.setCellStyle(estilo);
					      
						  cantTotalAnio03 = obtenerCantidadReclamosAnio(listaReporteAnioPasado, Constantes.ID_ESTADO_RECLAMO_FILE_NO_EXISTENTE);
						  celdaExcel = filaExcel.createCell((short)1);
					      totalAnioPasado+=objReporteAnioPasadoj.getTotales();
					      celdaExcel.setCellValue(cantTotalAnio03);
					      celdaExcel.setCellStyle(estilo);
						  for (String string : listaMes) {
							  cantTotalMes03 = obtenerCantidadReclamosMes(listaMatriz.get(string) , Constantes.ID_ESTADO_RECLAMO_FILE_NO_EXISTENTE);
							  celdaExcel = filaExcel.createCell((short)num);
							  strTemporal = String.valueOf(cantTotalMes03);
							  totalAnioActualNoExiste+=cantTotalMes03;
							  celdaExcel.setCellValue(strTemporal);
							  celdaExcel.setCellStyle(estilo);
					          num ++;
					          if(string.equals("ENERO")){
					        	totalReporte01=totalReporte01+cantTotalMes03;
					          }else if (string.equals("FEBRERO")){
					        	totalReporte02=totalReporte02+cantTotalMes03;
					          }else if (string.equals("MARZO")){
					        	totalReporte03=totalReporte03+cantTotalMes03;
					          }else if (string.equals("ABRIL")){
					        	totalReporte04=totalReporte04+cantTotalMes03;
					          }else if (string.equals("MAYO")){
					        	totalReporte05=totalReporte05+cantTotalMes03;
					          }else if (string.equals("JUNIO")){
					        	totalReporte06=totalReporte06+cantTotalMes03;
					          }else if (string.equals("JULIO")){
					        	totalReporte07=totalReporte07+cantTotalMes03;
					          }else if (string.equals("AGOSTO")){
					        	totalReporte08=totalReporte08+cantTotalMes03;
					          }else if (string.equals("SETIEMBRE")){
					        	totalReporte09=totalReporte09+cantTotalMes03;
					          }else if (string.equals("OCTUBRE")){
					        	totalReporte10=totalReporte10+cantTotalMes03;
					          }else if (string.equals("NOVIEMBRE")){
					        	totalReporte11=totalReporte11+cantTotalMes03;
					          }else if (string.equals("DICIEMBRE")){
					            	totalReporte12=totalReporte12+cantTotalMes03;
					          }
						  }
						  celdaExcel = filaExcel.createCell((short)num);
						  strTemporal = String.valueOf(totalAnioActualNoExiste);
						  celdaExcel.setCellValue(strTemporal);
						  celdaExcel.setCellStyle(estilo);
					  }
					  i++;
				  }
				  /***FIN DETALLE - EXCEL**/
				  filaExcel = hojaExcel.createRow(++numeroFila);
				  int contador=0;
				  int totalGeneralReporte=0;
				  totalGeneralReporte=totalAnioActualCerrado+totalAnioActualAbierto+totalAnioActualNoExiste+totalAnioActualError;
				  HSSFCellStyle estilo = ((contador%2)==0)?csCeldaFilaBlanca:csCeldaFilaColor;
				  String strTemporalTotal = "";  
				  short contadorCelda=0;
				  celdaExcel = filaExcel.createCell(contadorCelda);
				  strTemporalTotal = "TOTAL";
		          celdaExcel.setCellValue(strTemporalTotal);
		          celdaExcel.setCellStyle(estilo);
		          contadorCelda++;
	          
		          celdaExcel = filaExcel.createCell(contadorCelda);
		          strTemporalTotal = String.valueOf(totalAnioPasado);
		          celdaExcel.setCellValue(strTemporalTotal);
		          celdaExcel.setCellStyle(estilo);
		          contadorCelda++;
		          if (totalReporte01>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte01);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte02>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte02);
		              celdaExcel.setCellValue(strTemporalTotal);
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte03>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte03);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte03>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte03);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte04>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte04);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte05>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte05);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte06>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte06);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte07>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte07);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              contadorCelda++;
		          }
		          if (totalReporte08>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte08);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte09>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte09);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte10>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte10);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte11>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte11);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		              contadorCelda++;
		          }
		          if (totalReporte12>0){
		              celdaExcel = filaExcel.createCell(contadorCelda);
		   	          strTemporalTotal = String.valueOf(totalReporte12);
		              celdaExcel.setCellValue(strTemporalTotal);	
		              celdaExcel.setCellStyle(estilo);
		          }
		          contadorCelda--;
		          celdaExcel = filaExcel.createCell(contadorCelda);
		          strTemporalTotal = String.valueOf(totalGeneralReporte);
		          celdaExcel.setCellValue(strTemporalTotal);	
		          celdaExcel.setCellStyle(estilo);
			  }
			  /* Pie de pagina del reporte. */
			  numeroFila=(short)(numeroFila + 1);
			  filaExcel = hojaExcel.createRow(numeroFila);
			  celdaExcel = filaExcel.createCell((short)0);
			  celdaExcel.setCellStyle(uDocumentoExcel.generarEstiloLeyenda(HSSFCellStyle.ALIGN_CENTER, false));
			  celdaExcel.setCellValue("****FIN DE REPORTE****");
			  
			  /* Enviar archivo al navegador para su descarga. */
			  response.setContentType("application/vnd.ms-excel");
			  response.setHeader("Content-Disposition","attachment; filename=\"ReporteEstadoActual.xls\"");
			  OutputStream out = response.getOutputStream();
			  libroExcel.write(out);
			  out.flush();
			  out.close();
			  FacesContext.getCurrentInstance().responseComplete();
			  }else{
				  WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "valorFinal_ValorInicial"));
			  }
				
		   }else{
			   WebUtil.mensajeInformacion(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "reporteController_ValoresInicialFinal")); 
		   }
		  } catch (Exception exception) {
			LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(exception)[1], exception);
	    	String[] error = MensajeExceptionUtil.obtenerMensajeError(exception);
	    	String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
	    	WebUtil.mensajeError(mensajeError);
		  }
	  }

	  private List<Reclamo> obtenerListaReclamosHistorico(List<Reclamo> listaReporteAnioPasado) throws GmdException {
		List<Reclamo> listaReclamosHistorico = new ArrayList<Reclamo>();
		try {
			Map<String, Object> parametrosBusqueda = new HashMap<String, Object>();
			parametrosBusqueda.put("idGenerico", Constantes.ID_GENERICO_ESTADO_RECLAMO);
			parametrosBusqueda.put("stRegi", Constantes.IND_ESTADO_REGISTRO_ACTIVO);
			List<DetalleGenerico> listaEstadosReclamosGenerico = genericoService.listarDetalleGenerico(parametrosBusqueda);
			Reclamo reclamo;
			if (listaReporteAnioPasado == null || listaReporteAnioPasado.isEmpty() ) {
				for (DetalleGenerico detalleGenerico : listaEstadosReclamosGenerico) {
					reclamo = new Reclamo();
					reclamo.setEstadoReclamo(detalleGenerico.getId());
					reclamo.setDescripcionEstadoReclamo(detalleGenerico.getDescripcion());
					reclamo.setTotales(0);
					listaReclamosHistorico.add(reclamo);
				}
			} else {
				for (DetalleGenerico detalleGenerico : listaEstadosReclamosGenerico) {
					if (validarExisteReclamoHistorico(detalleGenerico.getId(),listaReporteAnioPasado)) {
						listaReclamosHistorico.add(obtenerReclamo(detalleGenerico.getId(),listaReporteAnioPasado));
					}else{
						reclamo = new Reclamo();
						reclamo.setEstadoReclamo(detalleGenerico.getId());
						reclamo.setDescripcionEstadoReclamo(detalleGenerico.getDescripcion());
						reclamo.setTotales(0);
						listaReclamosHistorico.add(reclamo);
					}
				}
			}
		}catch (Exception excepcion) {
			throw new GmdException(excepcion);
		}
		return listaReclamosHistorico;
	  }
	  private Reclamo obtenerReclamo(Integer id,List<Reclamo> listaReporteAnioPasado) throws GmdException {
		  Reclamo reclamotmp = new Reclamo();
		  try {
			  for (Reclamo reclamo : listaReporteAnioPasado) {
				  if (id.compareTo(reclamo.getEstadoReclamo()) == 0) {
					  reclamotmp = reclamo;
					  break;
				  }
			  }
		  }catch (Exception excepcion) {
			  throw new GmdException(excepcion);
		  }
		  return reclamotmp;
	  }
	  private boolean validarExisteReclamoHistorico(Integer idEstadoReclamo, List<Reclamo> listaReporteAnioPasado) throws GmdException {
		  boolean retorno = false;
		  try {
			  for (Reclamo reclamo : listaReporteAnioPasado) {
				  if (idEstadoReclamo.compareTo(reclamo.getEstadoReclamo()) == 0) {
					  retorno = true;
					  break;
				  }
			  }
		  }catch (Exception excepcion) {
			  throw new GmdException(excepcion);
		  }
		  return retorno;
	  }
	
	  private Integer obtenerCantidadReclamosMes(Object listaReclamoMes,Integer idEstadoReclamo) throws GmdException {
		  List<Reclamo> listaReclamos = new ArrayList<Reclamo>();
		  Integer cantReclamos = 0;
		  try {
			  listaReclamos = (List<Reclamo>) listaReclamoMes;
			  for (Reclamo reclamo : listaReclamos) {
				  if (reclamo.getEstadoReclamo().compareTo(idEstadoReclamo)==0) {
					  cantReclamos = reclamo.getTotales();
					  break;
				  }
			  }
		  }catch (Exception excepcion) {
			  throw new GmdException(excepcion);
		  }
		  return cantReclamos;
	  }
	  private Integer obtenerCantidadReclamosAnio(List<Reclamo> listaReporte, Integer idEstadoReclamo) throws GmdException {
		  Integer cantReclamos = 0;
		  try {
			  for (Reclamo reclamo : listaReporte) {
				  if (reclamo.getEstadoReclamo().compareTo(idEstadoReclamo)==0) {
					  cantReclamos = reclamo.getTotales();
					  break;
				  }
			  }
		  }catch (Exception excepcion) {
			  throw new GmdException(excepcion);
		  }
		  return cantReclamos;
	  }
	
	  public void obtenerListaEquipoTrabajo() {
		  try {
			  listaEquipoTrabajo = new ArrayList<EquipoTrabajo>();
			  Integer idFrenteAtencion = reclamo.getIdFrenteAtencion();
			  String descFrenteAtencion = reclamo.getDescripcionFrenteAtencion();
			  if (idFrenteAtencion != null && idFrenteAtencion > 0) {
				  FrenteAtencion frenteAtencion = new FrenteAtencion();
				  frenteAtencion.setId(idFrenteAtencion);
				  frenteAtencion.setDescripcion(descFrenteAtencion);
				  this.listaEquipoTrabajo = frenteAtencionService.listarEquipoTrabajo(frenteAtencion);
			  }
		  }catch (Exception excepcion) {
			  LOGGER.error(MensajeExceptionUtil.obtenerMensajeError(excepcion)[1], excepcion);
			  String[] error = MensajeExceptionUtil.obtenerMensajeError(excepcion);
			  String mensajeError = format(PropiedadesUtil.obtenerPropiedad(Constantes.ARCHIVO_MENSAJE, "sisrec_mensaje_error"), error[0]);
			  WebUtil.mensajeError(mensajeError);
		  }
	  }
	
	  public String getSeccionMes() {
		  return seccionMes;
	  }

	  public void setSeccionMes(String seccionMes) {
		  this.seccionMes = seccionMes;
	  }
		
	  public String getSeccionAnio() {
		  return seccionAnio;
	  }

	  public void setSeccionAnio(String seccionAnio) {
		  this.seccionAnio = seccionAnio;
	  }
	
	  public String getNombreMes() {
		  return nombreMes;
	  }
	  public void setNombreMes(String nombreMes) {
		  this.nombreMes = nombreMes;
	  }
	
	  public List<String> getListarSeccionesAnio() {
		  return listarSeccionesAnio;
	  }
	  public void setListarSeccionesAnio(List<String> listarSeccionesAnio) {
		  this.listarSeccionesAnio = listarSeccionesAnio;
	  }
	  public List<SelectItem> getListarSeccionesMes() {
		  return listarSeccionesMes;
	  }
	  public void setListarSeccionesMes(List<SelectItem> listarSeccionesMes) {
		  this.listarSeccionesMes = listarSeccionesMes;
	  }
	  public Map<String, String> getListaMeses() {
		  return listaMeses;
	  }
	  public void setListaMeses(Map<String, String> listaMeses) {
		  this.listaMeses = listaMeses;
	  }
	  public Map<String, String> getListaAnios() {
		  return listaAnios;
	  }
	  public void setListaAnios(Map<String, String> listaAnios) {
		  this.listaAnios = listaAnios;
	  }
	  public Reclamo getReclamo() {
		  return reclamo;
	  }
	  public void setReclamo(Reclamo reclamo) {
		  this.reclamo = reclamo;
	  }
	  public List<EquipoTrabajo> getListaEquipoTrabajo() {
		  return listaEquipoTrabajo;
	  }
	  public void setListaEquipoTrabajo(List<EquipoTrabajo> listaEquipoTrabajo) {
		  this.listaEquipoTrabajo = listaEquipoTrabajo;
	  }
	  public Integer getCantNroReclamosConsulta() {
		  return cantNroReclamosConsulta;
	  }
	  public void setCantNroReclamosConsulta(Integer cantNroReclamosConsulta) {
		  this.cantNroReclamosConsulta = cantNroReclamosConsulta;
	  }
	  
	  public HSSFSheet crearHojaExcelReporteCierreMes(HSSFWorkbook libroExcel, String nombreNuevaHoja, String periodo, Integer totalRegistros, String titulo) throws GmdException{
			
			try {
				int numeroHojas=0;
				HSSFSheet nuevaHoja = libroExcel.createSheet(nombreNuevaHoja);
				numeroHojas=libroExcel.getNumberOfSheets()-1;
				libroExcel.setSheetName(numeroHojas, nombreNuevaHoja+Integer.toString(libroExcel.getNumberOfSheets()));
				HSSFSheet hojaExcel = libroExcel.getSheetAt(numeroHojas);
				HSSFCell celda = null;

				  UDocumentoExcel uDocumentoExcel = new UDocumentoExcel(libroExcel);
				  HSSFCellStyle csCeldaFilaBlanca = uDocumentoExcel.generarEstiloCeldaTablaBlanco(HSSFCellStyle.ALIGN_LEFT);
				  HSSFCellStyle csCeldaFilaColor = uDocumentoExcel.generarEstiloCeldaTablaColor(HSSFCellStyle.ALIGN_LEFT);
				  HSSFCellStyle estiloCabeceraTabla = uDocumentoExcel.generarEstiloCabeceraTabla();
				  
				  
				  HSSFFont fuenteTitulo = libroExcel.createFont();
				  fuenteTitulo.setFontHeightInPoints((short) 11);
				  fuenteTitulo.setFontName(fuenteTitulo.FONT_ARIAL);
				  fuenteTitulo.setUnderline(HSSFFont.U_SINGLE);
				  fuenteTitulo.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				  
				  HSSFFont fuenteTituloGeneral = libroExcel.createFont();
				  fuenteTituloGeneral.setFontHeightInPoints((short) 11);
				  fuenteTituloGeneral.setFontName(fuenteTituloGeneral.FONT_ARIAL);
				  fuenteTituloGeneral.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
				  
				  HSSFCellStyle estiloTituloGeneral = libroExcel.createCellStyle();
					estiloTituloGeneral.setFont(fuenteTituloGeneral);
					estiloTituloGeneral.setAlignment(HSSFCellStyle.ALIGN_LEFT);
					estiloTituloGeneral.setVerticalAlignment(HSSFCellStyle.ALIGN_LEFT);
					estiloTituloGeneral.setWrapText(true);
					
					
					HSSFCellStyle estiloTitulo = libroExcel.createCellStyle();
					estiloTitulo.setFont(fuenteTitulo);
					estiloTitulo.setAlignment(HSSFCellStyle.ALIGN_CENTER);
					estiloTitulo.setVerticalAlignment(HSSFCellStyle.ALIGN_CENTER);
					estiloTitulo.setWrapText(true);
				  
				  HSSFRow filaExcel = null;
				  HSSFCell celdaExcel = null;
				  
				  /* Ancho por defecto de las columnas; importante no se configura el ancho de la celda 0 ya que esta viene de la plantilla excel por defecto. */
				  hojaExcel.setColumnWidth((short)(1),(short)4000);
		          hojaExcel.setColumnWidth((short)(2),(short)5000); 
		          hojaExcel.setColumnWidth((short)(3),(short)4000);
		          hojaExcel.setColumnWidth((short)(4),(short)8000);
		          hojaExcel.setColumnWidth((short)(5),(short)10000);
		          hojaExcel.setColumnWidth((short)(6),(short)5000);
		          hojaExcel.setColumnWidth((short)(7),(short)5000);
		          /*REQ001 - Item001 - Inicio*/
		          hojaExcel.setColumnWidth((short)(8),(short)5000);
		          /*REQ001 - Item001 - Fin*/
		          filaExcel = hojaExcel.createRow(4);
		          celdaExcel = filaExcel.createCell((short)0);

		          //Insertar el total de registros 
		          filaExcel = hojaExcel.createRow(5);
		          celdaExcel = filaExcel.createCell((short)0);
		          celdaExcel.setCellValue("Total Registros encontrados = " + totalRegistros);
		          //Insertar el mes y año del reporte
		          celdaExcel = filaExcel.createCell((short)4);
		          celdaExcel.setCellValue("Periodo: " + periodo);
		          
	        	  //Insertar la fecha
		          filaExcel = hojaExcel.createRow(1);
		          celda = filaExcel.createCell(6);
				  celda.setCellStyle(estiloTituloGeneral);
				  celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				  celda.setCellValue("Fecha: ");
		          celdaExcel = filaExcel.createCell((short)7);
		          Calendar fechaActual = Calendar.getInstance();
		          String strFechaActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.DAY_OF_MONTH)+"",2)+"/"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MONTH)+1)+"",2)+"/"+Funciones.CompletaCerosIzq(fechaActual.get(Calendar.YEAR)+"",4);
		          celdaExcel.setCellValue(strFechaActual);
				  
		          // Insertar la Hora
		          filaExcel = hojaExcel.createRow(2);
		          celda = filaExcel.createCell(3);
				  celda.setCellStyle(estiloTitulo);
				  celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				  celda.setCellValue(titulo);
				  hojaExcel.addMergedRegion(new Region(2, (short) 3, 2, (short) 4));
		          
		          celda = filaExcel.createCell(6);
				  celda.setCellStyle(estiloTituloGeneral);
				  celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				  celda.setCellValue("Hora: ");
		          celdaExcel = filaExcel.createCell((short)7);
		          String sHoraActual = Funciones.CompletaCerosIzq(fechaActual.get(Calendar.HOUR)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.MINUTE)+1)+"",2)+":"+Funciones.CompletaCerosIzq((fechaActual.get(Calendar.SECOND)+1)+"",2);
		          celdaExcel.setCellValue(sHoraActual);
		          
		          // Insertar el Usuario
		          filaExcel = hojaExcel.createRow(3);
		          celda = filaExcel.createCell(6);
				  celda.setCellStyle(estiloTituloGeneral);
				  celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				  celda.setCellValue("Usuario: ");  
		          celdaExcel = filaExcel.createCell((short)7);
		          celdaExcel.setCellValue(WebUtil.obtenerLoginUsuario());
		          
		        //Comenzando a llenar el XLS celda con titulos
		        filaExcel = hojaExcel.createRow(7);
				celda = filaExcel.createCell(0);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Item");
				
				celda = filaExcel.createCell(1);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Nro Reclamo");
				
				celda = filaExcel.createCell(2);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Estado");
				
				celda = filaExcel.createCell(3);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Frente");
				
				celda = filaExcel.createCell(4);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Equipo");
				
				celda = filaExcel.createCell(5);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Gestor");
				
				celda = filaExcel.createCell(6);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Mes Registro Reclamo");

				celda = filaExcel.createCell(7);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Fecha Registro Reclamo");
				
				/*REQ001 - Item001 - Inicio*/
				celda = filaExcel.createCell(8);
				celda.setCellStyle(estiloCabeceraTabla);
				celda.setCellType(HSSFCell.CELL_TYPE_STRING);
				celda.setCellValue("Fecha Modificación");
				/*REQ001 - Item001 -Fin*/
				return nuevaHoja;
			
			} catch (Exception excepcion) {
				throw new GmdException(excepcion);
			}
			
		}
	  
	  /*REQ001 - Item001 - Inicio*/
	  private List<Map<String, Object>> obtenerReclamosBaseDatos(
				List<Map<String, Object>> listaReclamosCierreMes,
				Integer ... estadoReclamo) throws GmdException {
	    	List<Map<String, Object>> listaReclamos = new ArrayList<Map<String,Object>>();
			try {
				if (listaReclamosCierreMes != null 
						&& !listaReclamosCierreMes.isEmpty()) {
					if (estadoReclamo.length > 1) {
						for (Map<String, Object> map : listaReclamosCierreMes) {
							for (int i = 0; i < estadoReclamo.length; i++) {
								if (estadoReclamo[i].compareTo(Integer.valueOf(map.get("Id Estado Reclamo").toString())) == 0) {
									listaReclamos.add(map);
								}
							}
						}
					} else {
						for (Map<String, Object> map : listaReclamosCierreMes) {
							if (estadoReclamo[0].compareTo(Integer.valueOf(map.get("Id Estado Reclamo").toString())) == 0) {
								listaReclamos.add(map);
							}
						}
					}
				}
			} catch (Exception excepcion) {
				throw new GmdException(excepcion);
			}
			return listaReclamos;
		}


	  /*REQ001 - Item001 - Fin*/
	public List<FrenteAtencion> obtenerListaFrenteAtencion() {
		List<FrenteAtencion> listaFrenteAtencionTmp = new ArrayList<FrenteAtencion>();
		try {
			listaFrenteAtencionTmp = frenteAtencionService.listar();
		}catch(Exception excepcion){
			String[] error = MensajeExceptionUtil.obtenerMensajeError(excepcion);
            LOGGER.error(error[1], excepcion);
		}
		return listaFrenteAtencionTmp;
	}
	
	public List<FrenteAtencion> getListaFrenteAtencion() {
		return listaFrenteAtencion;
	}

	public void setListaFrenteAtencion(List<FrenteAtencion> listaFrenteAtencion) {
		this.listaFrenteAtencion = listaFrenteAtencion;
	}

	public String getInicializar() {
		seccionMes = "";
		reclamo = new Reclamo();
		return inicializar;
	}

	public void setInicializar(String inicializar) {
		this.inicializar = inicializar;
	}
	
}
