
package com.mycompany.delacruz_crud_it1a;

import jxl.*;
import java.awt.*;
import java.io.*;
import java.sql.*;
import javax.swing.*;
import javax.swing.table.*;
import java.util.*;
import jxl.write.*;
import jxl.read.biff.BiffException;


public class ExcelImportExport extends javax.swing.JFrame {
    
    //Variable Declaration
    static JFileChooser jChooser;
    static Vector headers = new Vector();
    static DefaultTableModel model = null;
    static Vector data = new Vector();
    static int tblWidth = 0, tblHeight = 0;
    
    
    
    public ExcelImportExport(){
        initComponents();
        jChooser = new JFileChooser();
        model = new DefaultTableModel(data,headers);
        tbldata.setAutoCreateRowSorter(true);
        tbldata.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        tbldata.setRowHeight(25);
        tbldata.setRowMargin(4);
        tblWidth = model.getColumnCount()*150;
        tblHeight = model.getRowCount()*25;
        tbldata.setPreferredSize(new Dimension(tblWidth,tblHeight));
        
    }

    
    void fillData(File file)throws WriteException{
        Workbook workbook = null;
        try{
            try{
                workbook = Workbook.getWorkbook(file);
            }catch(IOException | BiffException e){
                JOptionPane.showMessageDialog(null, e);
            }
            
            Sheet sheet = workbook.getSheet(0); //sheet 1
            headers.clear();
            for(int i = 0;i<sheet.getColumns();i++){
                Cell cell1 = sheet.getCell(i,0);
                headers.add(cell1.getContents()); 
            }//end for loop
            data.clear();
            
            for (int j = 0; j<sheet.getRows();j++){
                Vector d = new Vector ();
                for(int i = 0;i<sheet.getColumns();i++){
                Cell cell = sheet.getCell(0,i);
                d.add(cell.getContents());
            }//end inner for loop
                d.add("\n");
                data.add(d);
            }//end outer for loop
            
        }catch(HeadlessException e){
                JOptionPane.showMessageDialog(null, e);
    }
    }
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        btnimport = new javax.swing.JButton();
        btnexport = new javax.swing.JButton();
        btnsave = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        tbldata = new javax.swing.JTable();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Import Export Excel File");
        setResizable(false);

        btnimport.setText("Import Excel File");
        btnimport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnimportActionPerformed(evt);
            }
        });

        btnexport.setText("Export to Excel File");
        btnexport.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btnexportActionPerformed(evt);
            }
        });

        btnsave.setText("Save Record");

        tbldata.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        tbldata.setPreferredSize(new java.awt.Dimension(300, 100));
        jScrollPane1.setViewportView(tbldata);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(51, 51, 51)
                .addComponent(btnimport)
                .addGap(79, 79, 79)
                .addComponent(btnexport)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                .addComponent(btnsave)
                .addGap(51, 51, 51))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                .addContainerGap(22, Short.MAX_VALUE)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 559, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(19, 19, 19))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(23, 23, 23)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(btnimport)
                    .addComponent(btnexport)
                    .addComponent(btnsave))
                .addGap(18, 18, 18)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 301, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(36, Short.MAX_VALUE))
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void btnexportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnexportActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_btnexportActionPerformed

    private void btnimportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btnimportActionPerformed
       jChooser.showOpenDialog(null);
       File file = jChooser.getSelectedFile();
       if(!file.getName().endsWith("xls")){
           JOptionPane.showMessageDialog(null, "Select .xls file only!", "File Extension Error",JOptionPane.ERROR_MESSAGE);
       }//end if
       else{
           try{
               fillData(file);
               model = new DefaultTableModel(data,headers);
               tblWidth = model.getColumnCount()*150;
               tblHeight = model.getRowCount()*25;
               tbldata.setPreferredSize(new Dimension(tblWidth,tblHeight));
               tbldata.setModel(model);
               
           }catch(WriteException e){
               JOptionPane.showMessageDialog(null, e);
           }
           
       }
    }//GEN-LAST:event_btnimportActionPerformed


    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ExcelImportExport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ExcelImportExport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ExcelImportExport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ExcelImportExport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ExcelImportExport().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btnexport;
    private javax.swing.JButton btnimport;
    private javax.swing.JButton btnsave;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable tbldata;
    // End of variables declaration//GEN-END:variables
}
