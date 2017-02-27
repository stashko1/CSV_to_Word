package CSV_to_Word;

import java.awt.Image;
import java.io.File;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileFilter;

/**
 *@author Stashko,Gurey,Miniajlenko,Vodolazskiy.
 * Класс в котором описано создание файла с расширением .docx для последующего
 * заполнения данніми с файла с расширением .csv. К этому классу поключена библиотека 
 * Swing, она служит для создания графического интерфейса для программ на языке Java. 
 */
public class Create_Doc extends javax.swing.JFrame {
    
    Find_and_read_csv obj = new Find_and_read_csv();
    /**
     * Creates new form Create_Doc
     */
    
    public Create_Doc() {
        Image icon = new ImageIcon(CSV_to_Word.Create_Doc.class.getResource("icon.png")).getImage();
        setTitle("CSV to DOCX");
        setIconImage(icon);
        initComponents();
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel1 = new javax.swing.JLabel();
        open_csv = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();
        open_directory = new javax.swing.JButton();
        Run = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("CSV to Word");
        setForeground(java.awt.Color.blue);

        jLabel1.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel1.setText("Выберите csv файл");

        open_csv.setFont(new java.awt.Font("Times New Roman", 0, 12)); // NOI18N
        open_csv.setText("Открыть");
        open_csv.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                open_csvActionPerformed(evt);
            }
        });

        jLabel2.setFont(new java.awt.Font("Times New Roman", 1, 14)); // NOI18N
        jLabel2.setText("Выберите место сохранения doc файла");

        open_directory.setFont(new java.awt.Font("Times New Roman", 0, 12)); // NOI18N
        open_directory.setText("Открыть");
        open_directory.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                open_directoryActionPerformed(evt);
            }
        });

        Run.setBackground(new java.awt.Color(204, 204, 204));
        Run.setFont(new java.awt.Font("Times New Roman", 3, 18)); // NOI18N
        Run.setForeground(new java.awt.Color(255, 0, 51));
        Run.setText("Запустить ");
        Run.setActionCommand("Запустить ");
        Run.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                RunActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(30, 30, 30)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(Run, javax.swing.GroupLayout.PREFERRED_SIZE, 338, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 30, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(jLabel1)
                            .addComponent(jLabel2))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 41, Short.MAX_VALUE)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(open_csv, javax.swing.GroupLayout.Alignment.TRAILING)
                            .addComponent(open_directory, javax.swing.GroupLayout.Alignment.TRAILING))))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(22, 22, 22)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel1)
                    .addComponent(open_csv))
                .addGap(18, 18, 18)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 58, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(open_directory))
                .addGap(30, 30, 30)
                .addComponent(Run, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addContainerGap(103, Short.MAX_VALUE))
        );

        jLabel1.getAccessibleContext().setAccessibleName("label_open");

        pack();
    }// </editor-fold>//GEN-END:initComponents
/* Метод для открытия файлов с расширением .csv. 
 * 
 * @param evt 
 */
    private void open_csvActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_open_csvActionPerformed
        JFileChooser opening_csv = new JFileChooser();
        FileFilter filter = new FileFilter() {
            private String extes = ".csv";
            public boolean accept(File f) {
                return (f.getName().endsWith(extes)); //To change body of generated methods, choose Tools | Templates.
            }

            @Override
            public String getDescription() {
                return "(*" + extes + ")"; //To change body of generated methods, choose Tools | Templates.
            }
        };
        opening_csv.setAcceptAllFileFilterUsed(false);
        opening_csv.setFileFilter(filter);
        File dir = new File("C:\\Users\\denrro\\Desktop\\CSV_to_Word\\data");
        opening_csv.setCurrentDirectory(dir);
        
        opening_csv.showOpenDialog(this);
        dir = opening_csv.getSelectedFile();
        obj.wave_to_csv = dir.getAbsolutePath();
    }//GEN-LAST:event_open_csvActionPerformed
/*
Метод для открытия проводника при поиске файла с расширением .csv.
К этому методу подключена библиотека java.awt.event, котрая служит для обработки
события.
    */
    private void open_directoryActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_open_directoryActionPerformed
        JFileChooser dialog = new JFileChooser();
        dialog.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        dialog.setApproveButtonText("Выбрать");//выбрать название для кнопки согласия
        dialog.setDialogTitle("Выберите директорию для сохранения");// выбрать название
        dialog.setDialogType(JFileChooser.OPEN_DIALOG);// выбрать тип диалога Open или Save
        dialog.setMultiSelectionEnabled(false); // Разрегить выбор нескольки файлов
        dialog.showOpenDialog(this);
        File file = dialog.getSelectedFile();
        obj.wave_to_saving_docx = file.getAbsolutePath();
    }//GEN-LAST:event_open_directoryActionPerformed
/*
Метод в котором віізвается метод finallyze для обработки файла.
    */
    private void RunActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_RunActionPerformed
        // TODO add your handling code here:
        obj.finallyze();
        System.exit(0);
    }//GEN-LAST:event_RunActionPerformed

    /**
     * @param args the command line arguments
     */
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
            java.util.logging.Logger.getLogger(Create_Doc.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Create_Doc.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Create_Doc.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Create_Doc.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Create_Doc().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton Run;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JButton open_csv;
    private javax.swing.JButton open_directory;
    // End of variables declaration//GEN-END:variables
}