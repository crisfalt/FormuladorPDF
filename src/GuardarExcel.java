
import com.itextpdf.text.html.simpleparser.CellWrapper;
import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JTable;
import jxl.CellView;
import jxl.Workbook;
import jxl.biff.DisplayFormat;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author WIN8
 */
public class GuardarExcel {
    
    
    private WritableWorkbook workbook;
    private WritableSheet sheet;
    //defines varios tipos de font
    private WritableFont titleFont = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
    private WritableFont subtitleFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
    private WritableFont headersFont = new WritableFont(WritableFont.ARIAL, 12, WritableFont.BOLD);
    private WritableFont textFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.NO_BOLD);
    private WritableFont totalFont = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
    private WritableFont textTotales = new WritableFont(WritableFont.ARIAL, 9, WritableFont.BOLD);
    //definimos un formato con el font de titulo
    WritableCellFormat titleFormat = new WritableCellFormat(titleFont);
  
    //el write de los totales
    WritableCellFormat totalFormat = new WritableCellFormat(textTotales);
    //definimos un formato con el font de subtitulo
    WritableCellFormat subtitleFormat = new WritableCellFormat(subtitleFont);

    //definimos un formato con el font de cabeceras, mas el fondo de color verde claro
    WritableCellFormat headerFormat = new WritableCellFormat(headersFont);
    //formato para numeros con negrita
    WritableCellFormat numberBold = new WritableCellFormat(new NumberFormat("###,###,##0.00"));
    //definimos un formato con el font para texto
    WritableCellFormat textFormat = new WritableCellFormat(textFont);

    //definimos un formato de tipo numerico y le asignamos el font de texto
    WritableCellFormat numberFormat = new WritableCellFormat(new NumberFormat("###,###,##0.00"));
    private JTable tabla;
    private File archivo;
    private DecimalFormat df = new DecimalFormat("###,###,##0.00");
    
    public GuardarExcel() {
             
    }
    
    public boolean cargarArchivo(  JTable tabla , File archivo  ) {
        this.tabla = tabla;
        this.archivo = archivo;
        
        try {
            workbook = Workbook.createWorkbook(archivo);
            headerFormat.setBackground(Colour.LIGHT_GREEN);
            sheet = workbook.createSheet("Hoja1", 0);
            titleFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
            subtitleFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
            subtitleFormat.setAlignment(Alignment.CENTRE);
            subtitleFormat.setShrinkToFit(false);
            totalFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
            totalFormat.setAlignment(Alignment.LEFT);
            headerFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
            textFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
            numberFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
            numberFormat.setAlignment(Alignment.RIGHT);
            numberBold.setBorder(Border.ALL, BorderLineStyle.THIN);
            numberBold.setAlignment(Alignment.RIGHT);
            numberBold.setFont(totalFont);
            return true;
        } catch (WriteException | IOException ex) {
//            Logger.getLogger(GuardarExcel.class.getName()).log(Level.SEVERE, null, ex);
            VentanaError.mostarError(ex.getMessage());
            return false;
        }
    }
    
    private boolean isNum(String str) {
        try {
            Long.parseLong(str);
            return true;
        }
        catch(Exception ex) {
            return false;
        }
    }
    
    public boolean isNumeric(String str) {
        try {
            df.parse(str);
            return true;
//            return (str.matches("[+-]?\\d*(\\.\\d+)?(\\,\\d+)?") && str.equals("")==false);
        } catch (ParseException ex) {
//            Logger.getLogger(GuardarExcel.class.getName()).log(Level.SEVERE, null, ex);
            return false;
        }
    }
    
    private boolean isImportant( final String textIN ) {
        if( textIN.equals("TOTAL CAPITULO") || textIN.equals("COSTO INDIRECTO") || textIN.equals("COSTO DIRECTO") || textIN.equals("TOTAL PRESUPUESTO") ) {
            return true;
        }
        else return false;
    }
    
    public void save() {
        int filas = tabla.getRowCount();
        int columnas = tabla.getColumnCount();
        int contador = 0;
        int filaTitulos = 0;
        for( int i = 0 ; i < filas ; i++ ){ 
            for( int j = 0 ; j < columnas ; j++ ) {
                try {
                    if( i > Globales.filaCodigo) { //lo que va debajo del codigo o encabezado
                        if(!tabla.getValueAt(i, j).toString().equals("")) { // verificar que no lea espacios en blanco del jtable a exportar
                            if( ( j == 3 || j == 4)) { // si es la columna 3 o 4 se ponen los valores en formato numero en excel
                                sheet.addCell(new jxl.write.Number(j,i,Double.parseDouble((df.parseObject(tabla.getValueAt(i, j).toString())).toString()),numberFormat));
                            }
                            else {
                                if( j == 0 ) { // para poner los codigos en excel 
                                    if( tabla.getValueAt(i, j).toString().toCharArray().length == 2 ) { // si el ccodigo solo tiene 2 numeros
                                        sheet.addCell(new Label(j,i,tabla.getValueAt(i, j).toString(),totalFormat));
                                        sheet.addCell(new Label(1,i,tabla.getValueAt(i, 1).toString(),totalFormat));
                                    }
                                    else {
                                        sheet.addCell(new Label(j,i,tabla.getValueAt(i, j).toString(),textFormat)); //pongo los codigos sin negrita
                                        sheet.addCell(new Label(2,i,tabla.getValueAt(i, 2).toString(),textFormat)); // y las unidades 
                                    }
                                }
                                else {
                                    if( j == 1 && tabla.getValueAt(i, 0).toString().toCharArray().length != 2 ) {
                                        if( isImportant(tabla.getValueAt(i, j).toString())) { // los textos importantes , total y asi
                                            sheet.addCell(new Label(1,i,tabla.getValueAt(i, 1).toString(),totalFormat));
                                        }
                                        else sheet.addCell(new Label(1,i,tabla.getValueAt(i, 1).toString(),textFormat)); //los textos de la columna 1
                                    }
                                    //aqui empienzan todas las formulas de los valores totales
                                    else {
                                        if( j == 5 ) {
                                            StringBuffer buf = new StringBuffer();
                                            buf.append("PRODUCTO(D"+(i+1)+":E"+(i+1));
                                            sheet.addCell(new Formula( j,i,buf.toString(),numberFormat) );
                                            if( tabla.getValueAt(i, 1).toString().equals("TOTAL CAPITULO") ) { //las sumas de los totales de capitulo
                                                int posicionInicial = 0;
                                                int posicionFinal = 0;
                                                List<Integer> temporal = Globales.superTotal.get(contador);
                                                posicionInicial = temporal.get(0);
                                                posicionFinal = temporal.get(temporal.size()-1);
                                                StringBuffer buf2 = new StringBuffer();
                                                buf2.append("SUMA(F"+(posicionInicial+1)+":F"+(posicionFinal+1));

                                                sheet.addCell(new Formula(j,i,buf2.toString(),numberBold));
                                                contador++;
                                            }
                                            if( tabla.getValueAt(i, 1).toString().equals("TOTAL PRESUPUESTO") ) {
                                                if( Globales.filaDirecto == 0 ) {
                                                    StringBuffer buf2 = new StringBuffer();
                                                    buf2.append("SUMA(");
                                                    for( int index = 0 ; index < Globales.posicionesTotal.size() ; index++ ) {
                                                        buf2.append("F"+(Globales.posicionesTotal.get(index)+1));
                                                        if(index < Globales.posicionesTotal.size()-1) {
                                                            buf2.append(",");
                                                        }
                                                    }
                                                    buf2.append(")");
                                                    sheet.addCell(new Formula(j,i,buf2.toString(),numberBold));
                                                }                   
                                                else {
                                                    int contarTotalesIndirectos = 0; //contar los totales que esten despues de la fila directo
                                                    if( Globales.filaDirecto != 0 ) {
                                                        StringBuffer buf2 = new StringBuffer();
                                                        buf2.append("SUMA(");
                                                        StringBuffer buf3 = new StringBuffer();
                                                        buf3.append("SUMA(");
                                                        for( int index = 0 ; index < Globales.posicionesTotal.size() ; index++ ) {
                                                            if( Globales.posicionesTotal.get(index) < Globales.filaDirecto ) {
                                                                buf2.append("F"+(Globales.posicionesTotal.get(index)+1));
                                                                if(index < Globales.posicionesTotal.size()-1) {
                                                                    buf2.append(",");
                                                                }
                                                            }                
                                                            else {
                                                                buf3.append("F"+(Globales.posicionesTotal.get(index)+1));
                                                                if(index < Globales.posicionesTotal.size()-1) {
                                                                    buf3.append(",");
                                                                }
                                                                contarTotalesIndirectos++;
                                                            }
                                                        }
                                                        buf2.append(")");
                                                        buf3.append(")");
                                                        sheet.addCell(new Formula(j,Globales.filaDirecto,buf2.toString(),numberBold));
                                                        sheet.addCell(new Formula(j,Globales.filaIndirecto,buf3.toString(),numberBold));
                                                        List<Integer> listaPorcentajes = Globales.superTotal.get(Globales.superTotal.size()-contarTotalesIndirectos);
                                                        for( int indice = 0 ; indice < listaPorcentajes.size() ; indice++ ) {
                                                            if( tabla.getValueAt(listaPorcentajes.get(indice), 1).equals("IVA ") || tabla.getValueAt(listaPorcentajes.get(indice), 1).equals("IVA  ")) {
                                                                StringBuffer bufIva = new StringBuffer();
                                                                bufIva.append("PRODUCTO(F"+(listaPorcentajes.get(indice))+",D"+(listaPorcentajes.get(indice)+1)+"/100");
                                                                sheet.addCell(new Formula(4,listaPorcentajes.get(indice),bufIva.toString(),numberFormat));
                                                                sheet.addCell(new Formula(5,listaPorcentajes.get(indice),bufIva.toString(),numberFormat));
                                                            }
                                                            else {
                                                                StringBuffer bufPercent = new StringBuffer();
                                                                bufPercent.append("PRODUCTO(F"+(Globales.filaDirecto+1)+",D"+(listaPorcentajes.get(indice)+1)+"/100");
                                                                sheet.addCell(new Formula(4,listaPorcentajes.get(indice),bufPercent.toString(),numberFormat));
                                                                sheet.addCell(new Formula(5,listaPorcentajes.get(indice),bufPercent.toString(),numberFormat));
                                                            }
                                                        }
                                                        //formula total presupuesto para en casos de costos directos
                                                        StringBuffer totalPre = new StringBuffer();
                                                        totalPre.append("SUMA(F"+(Globales.filaDirecto+1)+",F"+(Globales.filaIndirecto+1)+")");
                                                        sheet.addCell(new Formula(j,Globales.filaPresupuesto,totalPre.toString(),numberBold));
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else sheet.addCell(new Label(j,i,tabla.getValueAt(i, j).toString(),textFormat));
                    }
                    else { //del encabezado para arriba
                        sheet.addCell(new Label(j,i,tabla.getValueAt(i, j).toString(),subtitleFormat));
                    }              
                } catch (WriteException ex) {
//                    Logger.getLogger(GuardarExcel.class.getName()).log(Level.SEVERE, null, ex);
                        VentanaError.mostarError(ex.getMessage());
                } catch (ParseException ex) {
                    Logger.getLogger(GuardarExcel.class.getName()).log(Level.SEVERE, null, ex);
                }
            }
        }
        
        try {
            //Para organizar el ancho de las columnas en el excel
            CellView cevell = new CellView();
            cevell.setDimension(10);
            sheet.setColumnView(0, cevell);
            CellView cevell1 = new CellView();
            cevell1.setDimension(55);
            sheet.setColumnView(1, cevell1);
            CellView cevell2 = new CellView();
            cevell2.setDimension(5);
            sheet.setColumnView(2, cevell2);
            CellView cevell3 = new CellView();
            cevell3.setDimension(12);
            sheet.setColumnView(3, cevell3);
            CellView cevell4 = new CellView();
            cevell4.setDimension(16);
            sheet.setColumnView(4, cevell4);
            CellView cevell5 = new CellView();
            cevell5.setDimension(16);
            sheet.setColumnView(5, cevell5);
            workbook.write();
            workbook.close();
            VentanaError.mostarSucces("Se genero el Excel con formulas");
        } catch (IOException ex) {
            System.out.println("GuardarExcel::guardar::save::IOexception " + ex.toString());
        } catch (WriteException ex) {
            System.out.println("GuardarExcel::guardar::save::WriteException " + ex.toString());
        }
        
    }
    
}
