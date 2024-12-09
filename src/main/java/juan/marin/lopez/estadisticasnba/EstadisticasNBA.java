/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JPanel.java to edit this template
 */
package juan.marin.lopez.estadisticasnba;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFrame;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author ADMINISTRADORGS2
 */
public class EstadisticasNBA extends javax.swing.JPanel {

    /**
     * Creates new form InterfazEstadisticasNBA
     */
    public EstadisticasNBA() {
        initComponents();    
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {
        java.awt.GridBagConstraints gridBagConstraints;

        LabelNombreJugador = new javax.swing.JLabel();
        TextNombreJugador = new javax.swing.JTextField();
        Label2Lanzados = new javax.swing.JLabel();
        Label2Encestados = new javax.swing.JLabel();
        SpinnerTotales2 = new javax.swing.JSpinner();
        Spinner2Encestados = new javax.swing.JSpinner();
        Label3Totales = new javax.swing.JLabel();
        SpinnerTotales3 = new javax.swing.JSpinner();
        Label3Encestados = new javax.swing.JLabel();
        Spinner3Encestados = new javax.swing.JSpinner();
        ButtonGuardar = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        SpinnerTotalesLibres = new javax.swing.JSpinner();
        SpinnerLibresEncestados = new javax.swing.JSpinner();

        setLayout(new java.awt.GridBagLayout());

        LabelNombreJugador.setText("Nombre Jugador");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(LabelNombreJugador, gridBagConstraints);

        TextNombreJugador.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                TextNombreJugadorActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 4;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.gridwidth = 3;
        gridBagConstraints.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(TextNombreJugador, gridBagConstraints);

        Label2Lanzados.setText("Tiros de 2 lanzados");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 9;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(Label2Lanzados, gridBagConstraints);

        Label2Encestados.setText("Tiros de 2 encestados");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 11;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(Label2Encestados, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 5;
        gridBagConstraints.gridy = 9;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(SpinnerTotales2, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 5;
        gridBagConstraints.gridy = 11;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(Spinner2Encestados, gridBagConstraints);

        Label3Totales.setText("Tiros de 3 lanzados");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 13;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(Label3Totales, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 5;
        gridBagConstraints.gridy = 13;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(SpinnerTotales3, gridBagConstraints);

        Label3Encestados.setText("Tiros de 3 encestados");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 15;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(Label3Encestados, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 5;
        gridBagConstraints.gridy = 15;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(Spinner3Encestados, gridBagConstraints);

        ButtonGuardar.setText("Guardar");
        ButtonGuardar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                ButtonGuardarActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 4;
        gridBagConstraints.gridy = 17;
        gridBagConstraints.insets = new java.awt.Insets(15, 15, 15, 0);
        add(ButtonGuardar, gridBagConstraints);

        jLabel1.setText("Tiros Libres lanzados");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 5;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(jLabel1, gridBagConstraints);

        jLabel2.setText("Tiros libres encestados");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 7;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(jLabel2, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 5;
        gridBagConstraints.gridy = 5;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gridBagConstraints.ipadx = 17;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(SpinnerTotalesLibres, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 5;
        gridBagConstraints.gridy = 7;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.HORIZONTAL;
        gridBagConstraints.insets = new java.awt.Insets(30, 30, 30, 30);
        add(SpinnerLibresEncestados, gridBagConstraints);
    }// </editor-fold>//GEN-END:initComponents

    private void TextNombreJugadorActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_TextNombreJugadorActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_TextNombreJugadorActionPerformed

    private void ButtonGuardarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_ButtonGuardarActionPerformed
        // TODO add your handling code here:
        
        String nombreJugador = TextNombreJugador.getText();
        int tirosLibresRealizados = (Integer) SpinnerTotalesLibres.getValue();
        int tirosLibresEncestados = (Integer) SpinnerLibresEncestados.getValue();
        int tiros2Realizados = (Integer) SpinnerTotales2.getValue();
        int tiros2Encestados = (Integer) Spinner2Encestados.getValue();
        int tiros3Realizados = (Integer) SpinnerTotales3.getValue();
        int tiros3Encestados = (Integer) Spinner3Encestados.getValue();
        
        // Calcular los puntos
        int puntos = (2 * tiros2Encestados) + (3 * tiros3Encestados) + (tirosLibresEncestados);

        // Calcular FGA y FTA
        int fga = tiros2Realizados + tiros3Realizados;
        int fta = tirosLibresRealizados;
        
        // Calcular TS% (True Shooting Percentage)
        double ts = (double) puntos / (2 * (fga + 0.44 * fta));

        // Calcular FG% (Field Goal Percentage)
        double fg = (fga > 0) ? (double)(tiros2Encestados + tiros3Encestados) / fga * 100 : 0;

        // Calcular eFG% (Effective Field Goal Percentage)
        double efg = (fga > 0) ? (double)(tiros2Encestados + (0.5 * tiros3Encestados)) / fga * 100 : 0;
        
        int tirosTotalesRealizados = tiros2Realizados + tiros3Realizados + tirosLibresRealizados;
        int tirosTotalesEncestados = tiros2Encestados + tiros3Encestados + tirosLibresEncestados;
        
        
        try {
        // Cargar archivo Excel existente, o crear uno si no existe
            File archivo = new File("EstadisticasNBA.xlsx");
            Workbook wb;
            Sheet sheet;

            if (archivo.exists()) {
                // Si el archivo existe, abrirlo
                FileInputStream fis = new FileInputStream(archivo);
                wb = new XSSFWorkbook(fis);
                sheet = wb.getSheetAt(0);
                fis.close();
            } else {
                // Si no existe, crearlo desde cero
                wb = new XSSFWorkbook();
                sheet = wb.createSheet("Estadísticas");

                // Crear las cabeceras (agregar FG%, eFG%, Tiros Totales)
                Row cabecera = sheet.createRow(0);
                cabecera.createCell(0).setCellValue("Nombre Jugador");
                cabecera.createCell(1).setCellValue("Tiros Totales Realizados");
                cabecera.createCell(2).setCellValue("Tiros Libres Realizados");
                cabecera.createCell(3).setCellValue("Tiros Libres Encestados");
                cabecera.createCell(4).setCellValue("Tiros 2 Realizados");
                cabecera.createCell(5).setCellValue("Tiros 2 Encestados");
                cabecera.createCell(6).setCellValue("Tiros 3 Realizados");
                cabecera.createCell(7).setCellValue("Tiros 3 Encestados");
                cabecera.createCell(8).setCellValue("Tiros Totales Encestados");
                cabecera.createCell(9).setCellValue("% FG");
                cabecera.createCell(10).setCellValue("% eFG");
                cabecera.createCell(11).setCellValue("% TS");
            }
               
            // Eliminar la fila de promedios si existe
            int rowCount = sheet.getPhysicalNumberOfRows();
            if (rowCount > 1) {
                Row lastRow = sheet.getRow(rowCount - 1);
                if (lastRow != null && lastRow.getCell(0).getStringCellValue().equals("Promedio")) {
                    sheet.removeRow(lastRow);
                }
            }
            
            // Crear una nueva fila para el nuevo jugador en la primera fila libre
            Row row = sheet.createRow(rowCount);
            row.createCell(0).setCellValue(nombreJugador);
            row.createCell(1).setCellValue(tirosTotalesRealizados);  // Agregar Tiros Totales Realizados
            row.createCell(2).setCellValue(tirosLibresRealizados);
            row.createCell(3).setCellValue(tirosLibresEncestados);
            row.createCell(4).setCellValue(tiros2Realizados);
            row.createCell(5).setCellValue(tiros2Encestados);
            row.createCell(6).setCellValue(tiros3Realizados);
            row.createCell(7).setCellValue(tiros3Encestados);
            row.createCell(8).setCellValue(tirosTotalesEncestados);  // Agregar Tiros Totales Encestados
            row.createCell(9).setCellValue(fg);  // Agregar % FG al archivo
            row.createCell(10).setCellValue(efg);  // Agregar % eFG al archivo
            row.createCell(11).setCellValue(ts);  // Agregar % TS al archivo
            
            
            
            // Calcular promedios
            double totalTirosTotalesRealizados = 0;
            double totalTiros2Realizados = 0;
            double totalTiros2Encestados = 0;
            double totalTiros3Realizados = 0;
            double totalTiros3Encestados = 0;
            double totalTirosLibresRealizados = 0;
            double totalTirosLibresEncestados = 0;
            double totalPuntos = 0;
            double totalFga = 0;
            double totalFta = 0;
            
            int numJugadores = sheet.getPhysicalNumberOfRows() - 1; // Excluir la fila de cabecera
            
            FileOutputStream fos = new FileOutputStream(archivo);
            wb.write(fos);
            
            for (int i = 1; i <= numJugadores; i++) {
                Row r = sheet.getRow(i);
                if (r == null) { // Verifica si la fila está vacía
                    sheet.removeRow(row); // Elimina la fila
                    i--;
                    return;
                }
                totalTirosTotalesRealizados += r.getCell(1).getNumericCellValue();
                totalTirosLibresRealizados += r.getCell(2).getNumericCellValue();
                totalTirosLibresEncestados += r.getCell(3).getNumericCellValue();
                totalTiros2Realizados += r.getCell(4).getNumericCellValue();
                totalTiros2Encestados += r.getCell(5).getNumericCellValue();
                totalTiros3Realizados += r.getCell(6).getNumericCellValue();
                totalTiros3Encestados += r.getCell(7).getNumericCellValue();
                totalPuntos += r.getCell(8).getNumericCellValue();
                totalFga += r.getCell(4).getNumericCellValue() + r.getCell(6).getNumericCellValue();
                totalFta += r.getCell(2).getNumericCellValue();
            }
            
            double promedioTirosTotalesRealizados = totalTirosTotalesRealizados / numJugadores;
            double promedioTirosLibresRealizados = totalTirosLibresRealizados / numJugadores;
            double promedioTirosLibresEncestados = totalTirosLibresEncestados / numJugadores;
            double promedioTiros2Realizados = totalTiros2Realizados / numJugadores;
            double promedioTiros2Encestados = totalTiros2Encestados / numJugadores;
            double promedioTiros3Realizados = totalTiros3Realizados / numJugadores;
            double promedioTiros3Encestados = totalTiros3Encestados / numJugadores;
            double promedioPuntos = totalPuntos / numJugadores;
            double promedioFg = (totalFga > 0) ? (totalTiros2Encestados + totalTiros3Encestados) / totalFga * 100 : 0;
            double promedioEfg = (totalFga > 0) ? (totalTiros2Encestados + 0.5 * totalTiros3Encestados) / totalFga * 100 : 0;
            double promedioTs = (totalFta > 0) ? totalPuntos/(2*(totalFga+0.44*totalFta)) : 0;

            // Agregar la fila de promedios
            Row promedioRow = sheet.createRow(rowCount+1);
            promedioRow.createCell(0).setCellValue("Promedio");
            promedioRow.createCell(1).setCellValue(promedioTirosTotalesRealizados);  // Agregar Tiros Totales Realizados
            promedioRow.createCell(2).setCellValue(promedioTirosLibresRealizados);
            promedioRow.createCell(3).setCellValue(promedioTirosLibresEncestados);
            promedioRow.createCell(4).setCellValue(promedioTiros2Realizados);
            promedioRow.createCell(5).setCellValue(promedioTiros2Encestados);
            promedioRow.createCell(6).setCellValue(promedioTiros3Realizados);
            promedioRow.createCell(7).setCellValue(promedioTiros3Encestados);
            promedioRow.createCell(8).setCellValue(promedioPuntos);  // Agregar Tiros Totales Encestados
            promedioRow.createCell(9).setCellValue(promedioFg);  // Agregar % FG al archivo
            promedioRow.createCell(10).setCellValue(promedioEfg);  // Agregar % eFG al archivo
            promedioRow.createCell(11).setCellValue(promedioTs);  // Agregar % TS al archivo
            
            fos = new FileOutputStream(archivo);
            wb.write(fos);
            fos.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }//GEN-LAST:event_ButtonGuardarActionPerformed

    public static void main(String args[]) {
        JFrame interfaz = new JFrame("Interfaz");
        EstadisticasNBA panel = new EstadisticasNBA();
        interfaz.setContentPane(panel);
        interfaz.setSize(485, 673);
        interfaz.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        interfaz.setLocationRelativeTo(null);
        interfaz.setVisible(true);
       
    }
    
    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton ButtonGuardar;
    private javax.swing.JLabel Label2Encestados;
    private javax.swing.JLabel Label2Lanzados;
    private javax.swing.JLabel Label3Encestados;
    private javax.swing.JLabel Label3Totales;
    private javax.swing.JLabel LabelNombreJugador;
    private javax.swing.JSpinner Spinner2Encestados;
    private javax.swing.JSpinner Spinner3Encestados;
    private javax.swing.JSpinner SpinnerLibresEncestados;
    private javax.swing.JSpinner SpinnerTotales2;
    private javax.swing.JSpinner SpinnerTotales3;
    private javax.swing.JSpinner SpinnerTotalesLibres;
    private javax.swing.JTextField TextNombreJugador;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    // End of variables declaration//GEN-END:variables
}
