package CSV_to_Word;
import java.io.FileInputStream;
import java.io.*;
import javax.swing.JOptionPane;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.util.ArrayList;
import java.nio.charset.*;
import java.nio.file.*;
import java.nio.file.Files.*;

/**
 *@author Stashko,Gurey,Miniajlenko,Vodolazskiy.
 *Класс в котором описаная основная задача программы, то есть : считывание данных с файла,
 * их форматирование, запиь данных в документ формата .docx.
 */
public class Find_and_read_csv 
{
    /*
    Конструктор без параметров.
    */
    Find_and_read_csv(){};

    /**
     *Первое поле, используется для считывания данных с файла формата .csv. 
     */
    public String wave_to_csv; 

    /**
     *Второе поле, используется для записи данных в файл формата .docx.
     */
    public String wave_to_saving_docx; 
    private ArrayList<String> write_  = new ArrayList<String>(); /* Инициализация 
    типизированной коллекции типа String*/
    /* 
    * Третье поле, используется для для очистки строк. 
    */
    private String s; 
    /* 
    Метод в котором мы описываем работу функции под названием UTF, котора я будет использоваться в дальнейшем.  
    @param s
    @exception JOptionPane класс для создания окон.
    */
    private void UTF(String wave_to_csv1){
        try{
            byte[] fileBytes = Files.readAllBytes(Paths.get(wave_to_csv1));
            s = new String(fileBytes, StandardCharsets.UTF_8);
            }
         catch(Exception e){
            JOptionPane.showMessageDialog(null, e);
        }
    }
    /* 
    Метод для чтения данных с файла (чтение выполняется за 1 цикл.)
    @param wave_to_csv
    @exception JOptionPane Класс для создания окон.
    */
    private void Reading(String wave_to_csv){
        try{
            UTF(wave_to_csv);
        }
        catch(Exception e){
            JOptionPane.showMessageDialog(null, e);
        }
    }
    /*
    * Метод с идентификатором ститик для очистки строк.
    @param s
    @param pos
    @return Возвращает результатирующую строку.
    */
    private static String removeCharAt(String s, int pos) {

       return s.substring(0,pos)+s.substring(pos+1);

    }
    /* Метод с идентификатором ститик для очистки строк.
    @param s1
    @exception <E>
    */
    private void cutting(String s1){
        for(int i = 0; i< s1.lastIndexOf(",");i++){
            for(int a = 0; a<6;a++){
            write_.add(s1.substring(0, s1.indexOf(",")));
            s1=s1.replace(write_.get(write_.size()-1), "");
            s1=removeCharAt(s1, 0);
            }
            try{
            write_.add(s1.substring(0, s1.indexOf("\n")));
            }
            catch(Exception e){
                write_.add(s1);
            }
            if(s1=="")
                break;
            try{
                for(int c=0;s1.charAt(0)!='\n';){
                    if(s1!="")
                        s1 = removeCharAt(s1, c);
                    if(s1=="")
                        break;
                }
            }
            catch(Exception e){};
        }
    }

    /**
     *Финальный метод, который выйполняет всю работу, считывает. обрабатывет и 
     * записывает данные в файл формата .docx.
     *@param wave_to_csv
     *@param wave_to_saving_docx 
     *@param s
     *@param wave_to_saving_docx
     * @param write_
     */
    public void finallyze(){
        Reading(wave_to_csv);
        cutting(s);
        Outloading(wave_to_saving_docx, write_);
    }

    /**
     * Метод для записи данных в файл формата .docx.
     * @param wave_forCreatingDocx
     * @param data
     */
    public void Outloading(String wave_forCreatingDocx, ArrayList<String> data){
    try {
            wave_forCreatingDocx = wave_forCreatingDocx+"\\account.docx";
            FileOutputStream outStream = new FileOutputStream(wave_forCreatingDocx);
            
            XWPFDocument doc = new XWPFDocument();
            XWPFParagraph begin = doc.createParagraph();
            begin.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun beginRun = begin.createRun();
            beginRun.setFontFamily("Comic Sans MS");
            beginRun.setFontSize(15);
            beginRun.setText("Information about accounts");
            ArrayList<XWPFParagraph> paragraphes = new ArrayList<XWPFParagraph>();
            paragraphes.add (doc.createParagraph());
            paragraphes.get(0).setAlignment(ParagraphAlignment.LEFT);
            
            for(int i = 0; i<write_.size(); i++){
                paragraphes.add (doc.createParagraph());
                XWPFRun paraRun = paragraphes.get(i).createRun();
                paraRun.setFontFamily("Comic Sans MS");
                paraRun.setFontSize(12);
                paraRun.setText("First Name: "+write_.get(i));
                
                paragraphes.add (doc.createParagraph());
                XWPFRun paraRun1 = paragraphes.get(++i).createRun();
                paraRun1.setFontFamily("Comic Sans MS");
                paraRun1.setFontSize(12);
                paraRun1.setText("Last Name: "+write_.get(i));
                
                paragraphes.add (doc.createParagraph());
                XWPFRun paraRun2 = paragraphes.get(++i).createRun();
                paraRun2.setFontFamily("Comic Sans MS");
                paraRun2.setFontSize(12);
                paraRun2.setText("Email: "+write_.get(i));
                
                paragraphes.add (doc.createParagraph());
                XWPFRun paraRun3 = paragraphes.get(++i).createRun();
                paraRun3.setFontFamily("Comic Sans MS");
                paraRun3.setFontSize(12);
                paraRun3.setText("Password: "+write_.get(i));
                
                paragraphes.add (doc.createParagraph());
                XWPFRun paraRun4 = paragraphes.get(++i).createRun();
                paraRun4.setFontFamily("Comic Sans MS");
                paraRun4.setFontSize(12);
                paraRun4.setText("Secondary email: "+write_.get(i));
                
                paragraphes.add (doc.createParagraph());
                XWPFRun paraRun5 = paragraphes.get(++i).createRun();
                paraRun5.setFontFamily("Comic Sans MS");
                paraRun5.setFontSize(12);
                paraRun5.setText("Mobile Phone 1: "+write_.get(i));
                
                
                paragraphes.add (doc.createParagraph());
                XWPFRun paraRun6 = paragraphes.get(++i).createRun();
                paraRun6.setFontFamily("Comic Sans MS");
                paraRun6.setFontSize(12);
                paraRun6.setText("Department: "+write_.get(i));
                
            }
          
            doc.write(outStream);
            outStream.close();
            JOptionPane.showMessageDialog(null, "Успешно сохранено");
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, e);
        }
    }
}
