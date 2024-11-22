/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.mycompany.nba;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Sheet;


import javax.swing.*;

import java.io.FileOutputStream;
import java.io.IOException;
/**
 *
 * @author jmari
 */
public class NBA extends javax.swing.JFrame {

    public NBA() {
        initComponents();
        this.setLocationRelativeTo(null);
    }

  
    @SuppressWarnings("unchecked")                    
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        BotonAñadir = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jSpinnerEfectuados = new javax.swing.JSpinner();
        jSpinnerTiros3 = new javax.swing.JSpinner();
        jLabel4 = new javax.swing.JLabel();
        jSpinnerTiros2 = new javax.swing.JSpinner();
        Jugador = new javax.swing.JTextField();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(255, 0, 255));

        BotonAñadir.setText("Calcular");
        BotonAñadir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                BotonCalcularActionPerformed(evt);
            }
        });

        jLabel1.setForeground(new java.awt.Color(255, 255, 255));
        jLabel1.setText("Nombre del jugador");

        jLabel2.setForeground(new java.awt.Color(255, 255, 255));
        jLabel2.setText("Tiros Efectuados");

        jLabel3.setForeground(new java.awt.Color(255, 255, 255));
        jLabel3.setText("Tiros de 2 Encestados");

        jLabel4.setForeground(new java.awt.Color(255, 255, 255));
        jLabel4.setText("Tiros de 3 Encestados");

        Jugador.setBackground(new java.awt.Color(255, 255, 255));
        Jugador.setForeground(new java.awt.Color(0, 0, 0));
        Jugador.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                JugadorActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGap(39, 39, 39)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 112, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 48, Short.MAX_VALUE)
                        .addComponent(Jugador, javax.swing.GroupLayout.PREFERRED_SIZE, 150, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 112, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 112, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 112, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jSpinnerEfectuados, javax.swing.GroupLayout.DEFAULT_SIZE, 76, Short.MAX_VALUE)
                            .addComponent(jSpinnerTiros2)
                            .addComponent(jSpinnerTiros3))))
                .addGap(33, 33, 33))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(BotonAñadir, javax.swing.GroupLayout.PREFERRED_SIZE, 99, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(136, 136, 136))
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGap(36, 36, 36)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGap(1, 1, 1)
                        .addComponent(Jugador))
                    .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(32, 32, 32)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jSpinnerEfectuados, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabel2, javax.swing.GroupLayout.DEFAULT_SIZE, 25, Short.MAX_VALUE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 40, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel3, javax.swing.GroupLayout.DEFAULT_SIZE, 25, Short.MAX_VALUE)
                    .addComponent(jSpinnerTiros2, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(40, 40, 40)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel4, javax.swing.GroupLayout.PREFERRED_SIZE, 22, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jSpinnerTiros3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(38, 38, 38)
                .addComponent(BotonAñadir, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(43, 43, 43))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
        );

        pack();
    }                  
    
        
    private void BotonCalcularActionPerformed(java.awt.event.ActionEvent evt) {                                              
           procesarYGenerarExcel(); 
    }                                             

    private void JugadorActionPerformed(java.awt.event.ActionEvent evt) {                                        
    }                                       
    
     private void procesarYGenerarExcel() {
        try {
            String nombreJugador = Jugador.getText();
            int tirosRealizados = (int) jSpinnerEfectuados.getValue();
            int tirosMetidos2 = (int) jSpinnerTiros2.getValue();
            int tirosMetidos3 = (int) jSpinnerTiros3.getValue();
            
              if (nombreJugador.isEmpty()) {
                JOptionPane.showMessageDialog(null, "Nombre del jugador.");
                return;
            }

            if (tirosRealizados == 0) {
                JOptionPane.showMessageDialog(null, "No puede haber realizado 0 tiros.");
                return;
            }

            double porcentajeFG = ((double) (tirosMetidos2 + tirosMetidos3) / tirosRealizados) * 100;
            double porcentajeEFG = ((double) (tirosMetidos2 + 1.5 * tirosMetidos3) / tirosRealizados) * 100;

            crearInformeExcel("C:\\Users\\jmari\\Documents\\NetBeansProjects\\NBA\\puntosNBA.xlsx", nombreJugador, tirosRealizados, tirosMetidos2, tirosMetidos3, porcentajeFG, porcentajeEFG);

            JOptionPane.showMessageDialog(null, "Datos guardados de forma exitosa en: puntosNBA.xlsx");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Error al procesar los datos: " + ex.getMessage());
        }
    }
     
    private void crearInformeExcel(String nombreArchivo, String nombreJugador, int tirosRealizados, int tirosMetidos2, int tirosMetidos3, double porcentajeFG, double porcentajeEFG) throws IOException {
       Workbook libroExcel;
       Sheet hoja;

       File archivoExcel = new File(nombreArchivo);

       if (archivoExcel.exists()) {
           FileInputStream fis = new FileInputStream(archivoExcel);
           libroExcel = new XSSFWorkbook(fis);
           hoja = libroExcel.getSheetAt(0);
           fis.close();
       } else {
           libroExcel = new XSSFWorkbook();
           hoja = libroExcel.createSheet("EstadÃ­sticas");

           Row encabezado = hoja.createRow(0);
           encabezado.createCell(0).setCellValue("Nombre del Jugador");
           encabezado.createCell(1).setCellValue("Tiros Efectuados");
           encabezado.createCell(2).setCellValue("Tiros de 2 Encestados");
           encabezado.createCell(3).setCellValue("Tiros de 3 Encestados");
           encabezado.createCell(4).setCellValue("% FG");
           encabezado.createCell(5).setCellValue("% eFG");
       }

       int filaNueva = hoja.getLastRowNum() + 1;

       Row fila = hoja.createRow(filaNueva);
       fila.createCell(0).setCellValue(nombreJugador);
       fila.createCell(1).setCellValue(tirosRealizados);
       fila.createCell(2).setCellValue(tirosMetidos2);
       fila.createCell(3).setCellValue(tirosMetidos3);
       fila.createCell(4).setCellValue(porcentajeFG);
       fila.createCell(5).setCellValue(porcentajeEFG);

       for (int i = 0; i <= 5; i++) {
           hoja.autoSizeColumn(i);
       }

       try (FileOutputStream archivo = new FileOutputStream(nombreArchivo)) {
           libroExcel.write(archivo);
       }
       libroExcel.close();
    }
    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new NBA().setVisible(true);
            }
        });
    }
         
    private javax.swing.JButton BotonAñadir;
    private javax.swing.JTextField Jugador;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JSpinner jSpinnerEfectuados;
    private javax.swing.JSpinner jSpinnerTiros2;
    private javax.swing.JSpinner jSpinnerTiros3;                
}
