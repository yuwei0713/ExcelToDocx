package work;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.sql.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.jacob.activeX.*;
import com.jacob.com.*;

public class creatword {
	
	private boolean saveonexit;
	Dispatch doc = null;
	private ActiveXComponent word;
	private Dispatch documents;
	
	public static FileSystemView fsv = FileSystemView.getFileSystemView();
	private static JFileChooser filechooser;
	private static FileNameExtensionFilter access = new FileNameExtensionFilter(".accdb", "accdb");
	private static FileNameExtensionFilter xlsx = new FileNameExtensionFilter(".xlsx", "xlsx");
	private static FileNameExtensionFilter docx = new FileNameExtensionFilter(".docx", "docx");
	private static File files;
	public static File[] filesc;
	public static DefaultTableModel model;
	private Workbook workbook;
	private Sheet sheet;
	private static ArrayList title = new ArrayList();
	private static ArrayList exceldata = new ArrayList();
	private static ArrayList savedata = new ArrayList();
	private static ArrayList outputfile = new ArrayList();
	private static ArrayList outputfile_semester = new ArrayList();
	private static ArrayList outputfile_year = new ArrayList();
	private static ArrayList outputfile_teacher = new ArrayList();
	private static ArrayList excel_year = new ArrayList();
	private static ArrayList excel_semester = new ArrayList();
	private static ArrayList excel_number = new ArrayList();
	private static ArrayList repeat_output = new ArrayList();
	private static ArrayList number_sort = new ArrayList();
	private static ArrayList semester_sort = new ArrayList();
	private static ArrayList year_sort = new ArrayList();
	private static ArrayList name_sort = new ArrayList();
	private static ArrayList teacher_sort = new ArrayList();
	static String[] printdata = new String[14];
	static String filePath = fsv.getHomeDirectory().getAbsolutePath();
	static String datapath = "";
	static String excelpath = "";
	static String wordpath = "";
	static String word_path = "";
	static int teachers = 0;
	private static File file = fsv.getHomeDirectory();
	private static String path = fsv.getHomeDirectory().toString();
	private static String dir_path = path+"/教學成效檢討";
	private static List<String> list_combine  = new ArrayList<String>();
	
	
	static int excelrow;
	static int excelsheet;
	static int excelcolumn;
	static int which = -1;
	static String database_path;
	static Connection connDB;
	static ResultSetMetaData data = null;
	static PreparedStatement ps = null;
	static ResultSet rs;
	static Statement st;
	
	private static final Object[] classtitle = new Object[] {" ","課程名稱","課程代號","必選修","平均","學年","學期","開課學制","修習年級","授課老師","修課人數","不及格人數","成績低於20人數"};
	private static final Object[] studenttitle = new Object[] { " ", "學生姓名","學號","學年","學期","課程名稱","課程代號","班級","成績"};
	private static final Object[] outputtitle = new Object[] {"課程名稱","課程代號","必選修","平均","學年","學期","開課學制","修習年級","授課老師","修課人數","不及格人數","成績低於20人數"};
	private static DefaultTableModel modelfile;
	private static DefaultTableModel modelstudent;
	private static DefaultTableModel modelclass;
	private static DefaultTableModel modeloutput;
	private static JScrollPane filepane;
	private static JScrollPane classpane;
	private static JScrollPane studentpane;
	private static JScrollPane outputpane;
	private static JTable filetable;
	private static JTable classtable;
	private static JTable studenttable;
	private static JTable outputtable;
	private static int row = -1;
	
	
	public static void uniteDoc(List<String> fileList, String savepaths) {
		if(fileList == null || fileList.size() == 0){
			return;
		}
		int size = fileList.size();
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        app.setProperty("Visible", new Variant(false));
        Object docs = app.getProperty("Documents").toDispatch();
        Object doc = Dispatch.invoke(
        		(Dispatch) docs, 
        		"Open", 
        		Dispatch.Method, 
        		new Object[]{(String) fileList.get(size - 1),
        				new Variant(false),new Variant(true)}, 
        				new int[3]).toDispatch();
        for (int i = 0; i < fileList.size() - 1; i++) {  
            Dispatch.invoke(app.getProperty("Selection").toDispatch(),  
                "insertFile", Dispatch.Method, new Object[] {  
                        (String) fileList.get(i), "",  
                        new Variant(false), new Variant(false),  
                        new Variant(false) }, new int[3]); 
            Dispatch selection = Dispatch.get(app, "Selection").toDispatch();
            Dispatch.call(selection,  "InsertBreak" ,  new Variant(5) );
        }
        Dispatch.invoke((Dispatch) doc, "SaveAs", Dispatch.Method,  
            new Object[] { savepaths, new Variant(1) }, new int[3]);  
        Variant f = new Variant(false);  
        Dispatch.call((Dispatch) doc, "Close", f); 
        app.invoke("Quit", new Variant[] {}); 
	}

	public creatword()
	{
		if(word==null)
		{
			word = new ActiveXComponent("Word.Application");
			word.setProperty("Visible", new Variant(false));
		}
		if(documents==null)
		{
			documents = word.getProperty("Documents").toDispatch();
		}
		saveonexit = false;
	}
	public Dispatch open(String inputDoc)
	{
		return Dispatch.call(documents, "Open" , inputDoc).toDispatch();
	}
	public Dispatch select() {
        return word.getProperty("Selection").toDispatch();
    }
	public void movestart(Dispatch selection) {
        Dispatch.call(selection,"HomeKey",new Variant(6));
    }
	public boolean find(Dispatch selection,String toFindText) {
        Dispatch find = word.call(selection,"Find").toDispatch();
        Dispatch.put(find,"Text",toFindText);
        Dispatch.put(find,"Forward","True");
        Dispatch.put(find,"Format","True");
        Dispatch.put(find,"MatchCase","True");
        Dispatch.put(find,"MatchWholeWord","True");
        return Dispatch.call(find,"Execute").getBoolean();
    }
	public void replace(Dispatch selection,String newText) {
        Dispatch.put(selection,"Text",newText);
    }
	public void replaceall(Dispatch selection,String oldText,Object replaceObj) {
        movestart(selection);
        if(oldText.startsWith("table") || replaceObj instanceof ArrayList)
            replacetable(selection,oldText,(ArrayList) replaceObj);
        else {
            String newText = (String) replaceObj;
            if(newText==null)
                newText="";
            else{
                while(find(selection,oldText)) {
                    replace(selection,newText);
                    Dispatch.call(selection,"MoveRight");
                }
            }
        }
    }
	public void replacetable(Dispatch selection,String tableName,ArrayList dataList) {
        if(dataList.size() <= 1) {
            return;
        }
        String[] cols = (String[]) dataList.get(0);
        String tbIndex = tableName.substring(tableName.lastIndexOf("@") + 1);
        int fromRow = Integer.parseInt(tableName.substring(tableName.lastIndexOf("$") + 1,tableName.lastIndexOf("@")));
        Dispatch tables = Dispatch.get(doc,"Tables").toDispatch();
        Dispatch table = Dispatch.call(tables,"Item",new Variant(tbIndex)).toDispatch();
        Dispatch rows = Dispatch.get(table,"Rows").toDispatch();
        for(int i = 1;i < dataList.size();i ++) {
            String[] datas = (String[]) dataList.get(i);
            if(Dispatch.get(rows,"Count").getInt() < fromRow + i - 1) {
            	Dispatch.call(rows,"Add");
            }
            for(int j = 0;j < datas.length;j++) {
                Dispatch cell = Dispatch.call(table,"Cell",Integer.toString(fromRow + i - 1),cols[j]).toDispatch();
                Dispatch.call(cell,"Select");
                Dispatch font = Dispatch.get(selection,"Font").toDispatch();
                Dispatch.put(font,"Bold","0");
                Dispatch.put(font,"Italic","0");
                Dispatch.put(selection,"Text",datas[j]);
            }
        }
	}
	public void save(String outputPath) {
        Dispatch.call(Dispatch.call(word,"WordBasic").getDispatch(),"FileSaveAs",outputPath);
    }
	public void close(Dispatch doc) {
        Dispatch.call(doc,"Close",new Variant(saveonexit));
        word.invoke("Quit",new Variant[]{});
        word = null;
    }
	public void toWord(String inputPath,String outPath,HashMap data) {
       String oldText;
       Object newValue;
       try {
            if(doc==null)
            doc = open(inputPath);
            Dispatch selection = select();
            Iterator keys = data.keySet().iterator();
            while(keys.hasNext()) {
            	oldText = (String) keys.next();
                newValue = data.get(oldText);
                replaceall(selection,oldText,newValue);
            }
             save(outPath);
       } catch(Exception e) {
            e.printStackTrace();
       } finally {
            if(doc != null)
            	close(doc);
       }
	}
	
	public static void pushword()
	{
			HashMap data = new HashMap();
			data.put("$teacher$", printdata[0]);
			data.put("$year$", printdata[1]);
			data.put("$semester$", printdata[2]);
			data.put("$number$", printdata[3]);
			data.put("$name$", printdata[4]);
			data.put("$system$", printdata[5]);
			data.put("$grade$", printdata[6]);
			data.put("$type$", printdata[7]);
			data.put("$average$", printdata[8]);
			data.put("$people$", printdata[9] );
			data.put("$59$", printdata[10]);
			data.put("$20$", printdata[11]);
			data.put("$40$", printdata[12]);
			data.put("$fail$", printdata[13]);
			creatword jw2 = new creatword();
			word_path = dir_path+"/"+printdata[1]+printdata[2]+"_"+printdata[3]+"_"+printdata[4]+".doc";
			jw2.toWord(wordpath,word_path, data);
			teacher_sort(printdata[0],printdata[1],printdata[2],printdata[3],printdata[4]);
			//list_combine.add(word_path);
	}
	public static void teacher_sort(String teacher,String year,String semester,String number,String name) {
		//String teacher_compare;
		int flag = 0;
		if(number_sort.isEmpty()) {
			name_sort.add(name);
			teacher_sort.add(teacher);
			number_sort.add(number);
			year_sort.add(year);
			semester_sort.add(semester);
			teachers++;
		}else {
			for(int i=0;i<teacher_sort.size();i++) {
				if(teacher.equals(teacher_sort.get(i).toString())) {
					name_sort.add(i,name);
					teacher_sort.add(i,teacher);
					number_sort.add(i,number);
					year_sort.add(i,year);
					semester_sort.add(i,semester);
					break;
				}else {
					if(teachers==1) {
						name_sort.add(name);
						teacher_sort.add(teacher);
						number_sort.add(number);
						year_sort.add(year);
						semester_sort.add(semester);
						teachers++;
						break;
					}else {
						for(int j=0;j<teacher_sort.size();j++) {
							if(teacher.equals(teacher_sort.get(j).toString())){
								name_sort.add(j,name);
								teacher_sort.add(j,teacher);
								number_sort.add(j,number);
								year_sort.add(j,year);
								semester_sort.add(j,semester);
								flag = 1;
								break;
							}else if(j==teacher_sort.size()-1){
								name_sort.add(name);
								teacher_sort.add(teacher);
								number_sort.add(number);
								year_sort.add(year);
								semester_sort.add(semester);
								teachers++;
								flag = 1;
								break;
							}
						}
					}
				}
				if(flag == 1) {
					flag = 0;
					break;
				}
			}
		}
	}
	public static void data_sort(){
		String teacher = teacher_sort.get(0).toString();
		String next_teacher = "";
		String name = "";
		int number=0,year=0,k=0;
		String semester;
		for(int i=0;i<teacher_sort.size();i++)
		{
			if(!teacher.equals(teacher_sort.get(i))) {
				next_teacher = teacher_sort.get(i).toString();
				//學年排序
				/*for(int j=k;j<i-1;j++) {
					for(int l=k;l<i-j-1;l++){
						if(Integer.parseInt(year_sort.get(l).toString()) > Integer.parseInt(year_sort.get(l+1).toString())) {
							number = Integer.parseInt(number_sort.get(l).toString());
							year = Integer.parseInt(year_sort.get(l).toString());
							semester = semester_sort.get(l).toString();
							name = name_sort.get(l).toString();
							year_sort.add(l+2,year);
							number_sort.add(l+2,number);
							semester_sort.add(l+2,semester);
							year_sort.remove(l);
							number_sort.remove(l);
							semester_sort.remove(l);
						}
					}
				}*/
				for(int j=k;j<i-1;j++) {
					for(int l=k;l<i-(j-k)-1;l++){
						if(Integer.parseInt(number_sort.get(l).toString()) > Integer.parseInt(number_sort.get(l+1).toString())) {
							number = Integer.parseInt(number_sort.get(l).toString());
							year = Integer.parseInt(year_sort.get(l).toString());
							semester = semester_sort.get(l).toString();
							name = name_sort.get(l).toString();
							year_sort.add(l+2,year);
							number_sort.add(l+2,number);
							semester_sort.add(l+2,semester);
							name_sort.add(l+2,name);
							year_sort.remove(l);
							number_sort.remove(l);
							semester_sort.remove(l);
							name_sort.remove(l);
						}
					}
				}
				k=i;
				teacher = next_teacher;
				System.out.println(teacher);
			}else if(i==teacher_sort.size()-1) {
				for(int j=k;j<i;j++) {
					for(int l=k;l<i-(j-k);l++){
						if(Integer.parseInt(number_sort.get(l).toString()) > Integer.parseInt(number_sort.get(l+1).toString())) {
							number = Integer.parseInt(number_sort.get(l).toString());
							year = Integer.parseInt(year_sort.get(l).toString());
							semester = semester_sort.get(l).toString();
							name = name_sort.get(l).toString();
							year_sort.add(l+2,year);
							number_sort.add(l+2,number);
							semester_sort.add(l+2,semester);
							name_sort.add(l+2,name);
							year_sort.remove(l);
							number_sort.remove(l);
							semester_sort.remove(l);
							name_sort.remove(l);
						}
					}
				}
			}
		}
		for(int i=0;i<number_sort.size();i++) {
			word_path = dir_path+"/"+year_sort.get(i).toString()+semester_sort.get(i).toString()+"_"+number_sort.get(i).toString()+"_"+name_sort.get(i).toString()+".doc";
			list_combine.add(word_path);
		}
		
	}
	//判斷excel格式
	public static void readXls(String xlspath) throws IOException, SQLException
	{
		if (xlspath == null || !(WDWUtil.excelxls(xlspath) || WDWUtil.excelxlsx(xlspath)))  
        {  
			JOptionPane.showMessageDialog(null, "錯誤", "文件名不是excel格式",JOptionPane.ERROR_MESSAGE);
        }
	    InputStream is = new FileInputStream(xlspath); 
	    
	    Workbook wb = null;  

        if (WDWUtil.excelxls(xlspath))  
        {  
            wb = new HSSFWorkbook(is);  
        }  
        else  
        {  
            wb = new XSSFWorkbook(is);  
        }  
        excelsheet = wb.getNumberOfSheets();
        for (int numSheet = 0; numSheet < excelsheet; numSheet++) {
            org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(numSheet);
            if (sheet == null) {
                continue;
            }
            excelrow = sheet.getLastRowNum();
            excelcolumn = 13;
            //列橫行直
            //先列在行
            /*for (int cellNum = 0; cellNum < excelcolumn; cellNum++) {
               for (int rowNum = 0; rowNum < excelrow; rowNum++) {
            	   Row row = sheet.getRow(rowNum);
                   if (row == null) {
                       continue;
                   }
                Cell cell = row.getCell(cellNum);
                if (cell == null || !getValue(cell , rowNum , cellNum).startsWith("0")) {
                	continue;
                }
                }
            }*/
            
            
            //先行在列
           for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) {
                    continue;
                }
               for (int cellNum = 0; cellNum <= sheet.getColumnWidth(cellNum); cellNum++) {
                Cell cell0 = row.getCell(cellNum);
                if (cell0 == null || !getValue(cell0 , rowNum , cellNum).startsWith("0")) {   //不已96779开头的记录。  
                	continue;
                }
                }
            }
            database();
            exceldata.removeAll(exceldata);
        }
	}
	private static String getValue(Cell cell , int i , int cellNum) {
        String cellValue = "";
        int flag = 0;
		if (null != cell)  
        {
			if(i!=0)
			{
				switch(cell.getCellType()) {
	            	case STRING:
	            		cellValue = cell.getStringCellValue();  
	            		exceldata.add(cellValue.toString());
	            		break;
	            	case NUMERIC:
	            		cellValue = cell.getNumericCellValue() + "";
	            		exceldata.add(cellValue);
	                	break;
	            	case BOOLEAN:
	            	 	cellValue = cell.getBooleanCellValue() + "";
	            	 	exceldata.add(cellValue);
	            	 	break;
	            	/*case FORMULA:
	            	 	cellValue = cell.getCellFormula(); 
	                 	break;
	            	case BLANK:
	            		cellValue = "";  
	                	break; 
	            	case ERROR:
	            		cellValue = "非法字符";  
	            		break;
	            	default:  
	                	cellValue = "未知类型"; 
	                	break; */
	            	}
				/*if(cellNum == 0)
				{
					for(int num=0;i<excel_year.size();i++)
					{
						if(cellValue.toString().equals(excel_year.get(num).toString()))
						{
							flag = 1;
							break;
						}
					}
					if(flag==0)
					{
						excel_year.add(i);
					}
				}
				if(cellNum == 1)
				{
					for(int num=0;i<excel_semester.size();i++)
					{
						if(cellValue.toString().equals(excel_semester.get(num).toString()))
						{
							flag = 1;
							break;
						}
					}
					if(flag==0)
					{
						excel_semester.add(i);
					}
				}
				if(cellNum == 2)
				{
					for(int num=0;i<excel_number.size();i++)
					{
						if(cellValue.toString().equals(excel_number.get(num).toString()))
						{
							flag = 1;
							break;
						}
					}
					if(flag==0)
					{
						excel_number.add(i);
					}
				}*/
			}
        }
        return cellValue;
	}
	static class WDWUtil
	{
		 public static boolean excelxls(String filePath)  
		    {  
		        return filePath.matches("^.+\\.(?i)(xls)$");  
		    }  
		 public static boolean excelxlsx(String filePath)  
		    {  
		        return filePath.matches("^.+\\.(?i)(xlsx)$");  
		    }
	}
	
	public static void gui()
	{
		JPanel mainpage = new JPanel();
		mainpage.setLayout(null);
		
		JPanel datapage = new JPanel();
		datapage.setLayout(null);
		
		JPanel classpage = new JPanel();
		classpage.setLayout(null);
		
		JPanel studentpage = new JPanel();
		studentpage.setLayout(null);
		
		JPanel outputpage = new JPanel();
		outputpage.setLayout(null);
		
		JFrame main = new JFrame("教學成效處理");
		main.setSize(500,750);
		main.setLayout(new BorderLayout());
		main.add(mainpage , BorderLayout.CENTER);
		//main.add(datapage , BorderLayout.CENTER);
		//main.add(classpage , BorderLayout.CENTER);
		//main.add(studentpage , BorderLayout.CENTER);
		mainpage.setVisible(true);
		datapage.setVisible(false);
		classpage.setVisible(false);
		studentpage.setVisible(false);
		outputpage.setVisible(false);
		main.getContentPane().setBackground(Color.LIGHT_GRAY);
		
		
		JLabel title = new JLabel("教學成效處理");
		Font font = new Font(Font.DIALOG_INPUT, Font.ITALIC, 35);
		title.setFont(font);
		title.setBounds(140,20,250,80);
		mainpage.add(title);
		
		Font fontbutton = new Font(Font.DIALOG_INPUT, Font.BOLD, 18);
		JButton datainput = new JButton("資料匯入");
		datainput.setBounds(175, 180, 150, 50);
		datainput.setFont(fontbutton);
		mainpage.add(datainput);
		
		
		JButton findclass = new JButton("課程查詢");
		findclass.setBounds(175, 260, 150, 50);
		findclass.setFont(fontbutton);
		mainpage.add(findclass);
		
		
		JButton findstudent = new JButton("學生查詢");
		findstudent.setBounds(175, 330, 150, 50);
		findstudent.setFont(fontbutton);
		mainpage.add(findstudent);
		
		JButton dataoutput = new JButton("查看匯出");
		dataoutput.setBounds(175, 480, 150, 50);
		dataoutput.setFont(fontbutton);
		mainpage.add(dataoutput);
		
		JButton returnmain = new JButton("返回");
		JButton find = new JButton("查詢");
		JButton delete = new JButton("刪除");
		JButton modify = new JButton("修改");
		JButton add = new JButton("新增");
		JButton selectall = new JButton("全選");
		
		JButton output = new JButton("加入匯出");
		JButton count = new JButton("計算");
		JButton determine = new JButton("確定");
		JButton clearfile = new JButton("清空資料庫");
		JButton inputdatabase = new JButton("匯入資料庫");
		JButton inputexcel = new JButton("匯入Excel");
		JButton inputword = new JButton("匯入Word範例檔案");
		
		JMenuBar menubar = new JMenuBar();
		main.setJMenuBar(menubar);
		modelclass = new DefaultTableModel((Object[])classtitle , 0);
		final JTable classtable = new JTable(modelclass);
		classtable.getColumnModel().getColumn(0).setCellEditor(new DefaultCellEditor(new JCheckBox()));
		classtable.getColumnModel().getColumn(0).setPreferredWidth(30);
		classtable.getColumnModel().getColumn(0).setCellRenderer(new DefaultTableCellRenderer() {
	          private static final long serialVersionUID = 1L;
	          
	          public Component getTableCellRendererComponent(JTable jtable, Object obj, boolean flag, boolean flag1, int i, int j) {
	            JCheckBox checkBox = new JCheckBox();
	            if (obj instanceof Boolean) {
	              checkBox.setSelected(((Boolean)obj).booleanValue());
	            } else {
	              checkBox.setSelected(Boolean.FALSE.booleanValue());
	            } 
	            return checkBox;
	          }
	        });
		classtable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		classpane = new JScrollPane(classtable);
		classpane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		
		modelstudent = new DefaultTableModel((Object[])studenttitle, 0);
		final JTable studenttable = new JTable(modelstudent);
		studenttable.getColumnModel().getColumn(0).setCellEditor(new DefaultCellEditor(new JCheckBox()));
		studenttable.getColumnModel().getColumn(0).setPreferredWidth(30);
		studenttable.getColumnModel().getColumn(0).setCellRenderer(new DefaultTableCellRenderer() {
	          private static final long serialVersionUID = 1L;
	          
	          public Component getTableCellRendererComponent(JTable jtable, Object obj, boolean flag, boolean flag1, int i, int j) {
	            JCheckBox checkBox = new JCheckBox();
	            if (obj instanceof Boolean) {
	              checkBox.setSelected(((Boolean)obj).booleanValue());
	            } else {
	              checkBox.setSelected(Boolean.FALSE.booleanValue());
	            } 
	            return checkBox;
	          }
	        });
		//studenttable.setPreferredScrollableViewportSize(new Dimension(460,430));
		studenttable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		studentpane = new JScrollPane(studenttable);
		studentpane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		
		modeloutput = new DefaultTableModel((Object[])outputtitle, 0);
		final JTable outputtable = new JTable(modeloutput);
		outputtable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		outputpane = new JScrollPane(outputtable);
		outputpane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
		
		Font fontlabel = new Font(Font.DIALOG_INPUT, Font.BOLD, 18);
		
		//filepane.add(filetable);
		//studentpane.add(studenttable);
		//classpane.add(classtable);
		JTextField database_text = new JTextField();
		JTextField excel_text = new JTextField();
		JTextField word_text = new JTextField();
		
		JLabel classname = new JLabel("課程名稱");
		JLabel classnumber = new JLabel("課程代號");
		JLabel TakeCourse_Grade = new JLabel("修習年級");
		
		JTextField classname_text = new JTextField();
		JTextField classnumber_text = new JTextField();
		JTextField TakeCourse_Grade_text = new JTextField();
		
		classname.setFont(fontlabel);
		classnumber.setFont(fontlabel);
		TakeCourse_Grade.setFont(fontlabel);
		
		
		JLabel teachername = new JLabel("授課老師");
		JLabel SemesterType = new JLabel("學期");
		JLabel year = new JLabel("學年");
		JLabel School_System = new JLabel("開課學制");
		JLabel type = new JLabel("必選修");
		
		JLabel database = new JLabel("資料庫路徑");
		JLabel excel = new JLabel("Excel路徑");
		JLabel word = new JLabel("Word範例檔路徑");
		
		JTextField teachername_text = new JTextField();
		JComboBox SemesterType_box = new JComboBox();
		JTextField year_text = new JTextField();
		JComboBox School_System_box = new JComboBox();
		JComboBox type_box = new JComboBox();
		teachername.setFont(fontlabel);
		SemesterType.setFont(fontlabel);
		year.setFont(fontlabel);
		School_System.setFont(fontlabel);
		type.setFont(fontlabel);
		database.setFont(fontlabel);
		excel.setFont(fontlabel);
		word.setFont(fontlabel);
		
		type_box.setBackground(Color.WHITE);
		School_System_box.setBackground(Color.WHITE);
		SemesterType_box.setBackground(Color.WHITE);
		type_box.addItem("--");
		type_box.addItem("必修");
		type_box.addItem("必選");
		type_box.addItem("選修");
		School_System_box.addItem("--");
		School_System_box.addItem("日間部");
		School_System_box.addItem("碩士班");
		School_System_box.addItem("碩士在職專班");
		SemesterType_box.addItem("--");
		SemesterType_box.addItem("上");
		SemesterType_box.addItem("下");
		
		JLabel studentname = new JLabel("學生姓名");
		JLabel studentnumber = new JLabel("學號");
		JLabel grade = new JLabel("成績");
		
		JTextField studentname_text = new JTextField(60);
		JTextField studentnumber_text = new JTextField(13);
		JTextField grade_text = new JTextField(5);
		
		studentname.setFont(fontlabel);
		studentnumber.setFont(fontlabel);
		grade.setFont(fontlabel);
		
		datainput.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	which=0;
	        	main.setSize(500,550);
	        	main.add(datapage , BorderLayout.CENTER);
	        	mainpage.setVisible(false);
	    		datapage.setVisible(true);
	    		classpage.setVisible(false);
	    		studentpage.setVisible(false);
	    		
	    		determine.setBounds(420, 470, 60, 30);
	    		delete.setBounds(20, 470, 100, 30);
	    		database.setBounds(20, 35, 100, 40);
	    		excel.setBounds(20, 115, 100, 40);
	    		word.setBounds(20, 195, 160, 40);
	    		inputdatabase.setBounds(320, 79 , 160 , 30);
	    		inputexcel.setBounds(320, 159 , 160 , 30);
	    		inputword.setBounds(320, 239 , 160 , 30);
	    		database_text.setBounds(20, 80, 290, 30);
	    		excel_text.setBounds(20, 160, 290, 30);
	    		word_text.setBounds(20, 240, 290, 30);
	    		database_text.setText(datapath);
	    		
	    		datapage.add(database);
	    		datapage.add(excel);
	    		datapage.add(word);
	    		datapage.add(inputdatabase);
	    		datapage.add(inputexcel);
	    		datapage.add(inputword);
	    		datapage.add(determine);
	    		datapage.add(delete);
	    		datapage.add(database_text);
	    		datapage.add(excel_text);
	    		datapage.add(word_text);
	        }
	    });
		
		findclass.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	which=1;
	        	main.add(classpage , BorderLayout.CENTER);
	        	mainpage.setVisible(false);
	    		datapage.setVisible(false);
	    		classpage.setVisible(true);
	    		studentpage.setVisible(false);
	    		
	    		classname.setBounds(20 , 18 , 80 , 30);
	    		classname_text.setBounds(105 , 20 , 180 , 30);
	    		classnumber.setBounds(290 , 18 , 80 , 30);
	    		classnumber_text.setBounds(375 , 20 , 105 , 30);
	    		teachername.setBounds(20 , 58 , 80 , 30);
	    		teachername_text.setBounds(105 , 60 , 130 , 30);
	    		type.setBounds(310, 58, 60, 30);
	    		type_box.setBounds(375, 60, 70, 30);
	    		year.setBounds(20, 98, 40, 30);
	    		year_text.setBounds(65, 100, 75, 30);
	    		SemesterType.setBounds(150, 98, 40, 30);
	    		SemesterType_box.setBounds(195, 100, 70, 30);
	    		TakeCourse_Grade.setBounds(20, 138, 80, 30);
	    		TakeCourse_Grade_text.setBounds(105, 140, 130, 30);
	    		School_System.setBounds(290, 98, 80, 30);
	    		School_System_box.setBounds(375, 100, 105, 30);
	    		classpane.setBounds(20, 170, 460, 430);
	    		
	    		output.setBounds(130, 610, 100, 30);
	    		count.setBounds(20, 610, 100, 30);
	    		returnmain.setBounds(420, 660, 60, 30);
	    		delete.setBounds(20, 660, 100, 30);
	    		modify.setBounds(130, 660, 100, 30);
	    		add.setBounds(240, 660, 100, 30);
	    		find.setBounds(380, 610, 100, 30);
	    		selectall.setBounds(380 , 140, 100, 30);
	    		
	    		classpage.add(output);
	    		classpage.add(selectall);
	    		classpage.add(School_System_box);
	    		classpage.add(School_System);
	    		classpage.add(type);
	    		classpage.add(type_box);
	    		classpage.add(SemesterType);
	    		classpage.add(SemesterType_box);
	    		classpage.add(year);
	    		classpage.add(year_text);
	    		classpage.add(TakeCourse_Grade);
	    		classpage.add(TakeCourse_Grade_text);
	    		classpage.add(classpane);
	    		classpage.add(teachername);
	    		classpage.add(teachername_text);
	    		classpage.add(classnumber);
	    		classpage.add(classnumber_text);
	    		classpage.add(classname_text);
	    		classpage.add(classname);
	    		classpage.add(count);
	    		classpage.add(find);
	    		classpage.add(returnmain);
	    		classpage.add(delete);
	    		classpage.add(modify);
	    		classpage.add(add);
	    		find.doClick();
	        }
	    });
		
		findstudent.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	which=2;
	        	main.add(studentpage , BorderLayout.CENTER);
	        	mainpage.setVisible(false);
	    		datapage.setVisible(false);
	    		classpage.setVisible(false);
	    		studentpage.setVisible(true);
	    		
	    		studentname.setBounds(20, 18 , 80 , 30);
	    		studentname_text.setBounds(105, 20, 140, 30);
	    		studentnumber.setBounds(250, 18, 40, 30);
	    		studentnumber_text.setBounds(295, 20, 185, 30);
	    		classname.setBounds(20 , 58 , 80 , 30);
	    		classname_text.setBounds(105 , 60 , 180 , 30);
	    		classnumber.setBounds(290 , 58 , 80 , 30);
	    		classnumber_text.setBounds(375 , 60 , 105 , 30);
	    		TakeCourse_Grade.setBounds(20 , 98 , 80 , 30);
	    		TakeCourse_Grade_text.setBounds(105 , 100 , 180 , 30);
	    		grade.setBounds(360, 98, 40, 30);
	    		grade_text.setBounds(405, 100, 75, 30);
	    		year.setBounds(20, 138, 40, 30);
	    		year_text.setBounds(65, 140, 75, 30);
	    		SemesterType.setBounds(145, 138, 40, 30);
	    		SemesterType_box.setBounds(190, 140, 70, 30);
	    		
	    		studentpane.setBounds(20, 170, 460, 430);
	    		
	    		returnmain.setBounds(420, 660, 60, 30);
	    		delete.setBounds(20, 660, 100, 30);
	    		modify.setBounds(130, 660, 100, 30);
	    		add.setBounds(240, 660, 100, 30);
	    		find.setBounds(380, 610, 100, 30);
	    		selectall.setBounds(380 , 140, 100, 30);
	    		
	    		studentpage.add(year);
	    		studentpage.add(year_text);
	    		studentpage.add(SemesterType);
	    		studentpage.add(SemesterType_box);
	    		studentpage.add(selectall);
	    		studentpage.add(studentpane);
	    		studentpage.add(grade);
	    		studentpage.add(grade_text);
	    		studentpage.add(TakeCourse_Grade);
	    		studentpage.add(TakeCourse_Grade_text);
	    		studentpage.add(classname);
	    		studentpage.add(classname_text);
	    		studentpage.add(classnumber);
	    		studentpage.add(classnumber_text);
	    		studentpage.add(studentnumber);
	    		studentpage.add(studentnumber_text);
	    		studentpage.add(studentname);
	    		studentpage.add(studentname_text);
	    		studentpage.add(find);
	    		studentpage.add(returnmain);
	    		studentpage.add(delete);
	    		studentpage.add(modify);
	    		studentpage.add(add);
	    		find.doClick();
	        }
	    });
		
		dataoutput.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	which=10;
	        	main.add(outputpage , BorderLayout.CENTER);
	        	mainpage.setVisible(false);
	    		classpage.setVisible(false);
	    		studentpage.setVisible(false);
	    		outputpage.setVisible(true);
	    		
	    		determine.setBounds(420, 600, 60, 30);
	    		delete.setBounds(20, 600, 100, 30);
	    		returnmain.setBounds(420, 660, 60, 30);
	    		outputpane.setBounds(20, 50, 460, 530);
	    		
	    		outputpage.add(outputpane);
	    		outputpage.add(delete);
	    		outputpage.add(determine);
	    		outputpage.add(returnmain);
	    		find.doClick();
	        }
	    });
		
		classtable.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent e) {
					row = classtable.getSelectedRow();
       				classname_text.setText(classtable.getModel().getValueAt(row, 1).toString());
       				classnumber_text.setText(classtable.getModel().getValueAt(row, 2).toString());
       				teachername_text.setText(classtable.getModel().getValueAt(row, 9).toString());
       				year_text.setText(classtable.getModel().getValueAt(row, 5).toString());
       				TakeCourse_Grade_text.setText(classtable.getModel().getValueAt(row, 8).toString());
       				type_box.removeAllItems();
       				SemesterType_box.removeAllItems();
       				School_System_box.removeAllItems();
       				if(classtable.getModel().getValueAt(row, 3).toString().equals("必修"))
       				{
       					type_box.addItem("必修");
       					type_box.addItem("必選");
       					type_box.addItem("選修");
       					type_box.addItem("--");
       				}
       				else if(classtable.getModel().getValueAt(row, 3).toString().equals("選修"))
       				{
       					type_box.addItem("選修");
       					type_box.addItem("必修");
       					type_box.addItem("必選");
       					type_box.addItem("--");
       				}
       				else if(classtable.getModel().getValueAt(row, 3).toString().equals("必選"))
       				{
       					type_box.addItem("必選");
       					type_box.addItem("選修");
       					type_box.addItem("必修");
       					type_box.addItem("--");
       				}
       				
       				if(classtable.getModel().getValueAt(row, 7).toString().equals("日間部"))
       				{
       					School_System_box.addItem("日間部");
    					School_System_box.addItem("碩士班");
    					School_System_box.addItem("碩士在職專班");
    					School_System_box.addItem("--");
       				}
       				else if(classtable.getModel().getValueAt(row, 7).toString().equals("碩士班"))
       				{
    					School_System_box.addItem("碩士班");
    					School_System_box.addItem("日間部");
    					School_System_box.addItem("碩士在職專班");
    					School_System_box.addItem("--");
       				}
       				else if(classtable.getModel().getValueAt(row, 7).toString().equals("碩士在職專班"))
       				{
       					School_System_box.addItem("碩士在職專班");
       					School_System_box.addItem("日間部");
       					School_System_box.addItem("碩士班");
    					School_System_box.addItem("--");
       				}
       				
       				if(classtable.getModel().getValueAt(row, 6).toString().equals("上"))
       				{
    					SemesterType_box.addItem("上");
    					SemesterType_box.addItem("下");
    					SemesterType_box.addItem("--");
       				}
       				if(classtable.getModel().getValueAt(row, 6).toString().equals("下"))
       				{
    					SemesterType_box.addItem("下");
    					SemesterType_box.addItem("上");
    					SemesterType_box.addItem("--");
       				}
             }
		});
		
		studenttable.addMouseListener(new MouseAdapter() {
			public void mouseClicked(MouseEvent e) {
						row = studenttable.getSelectedRow();
						studentname_text.setText(studenttable.getModel().getValueAt(row, 1).toString());
						studentnumber_text.setText(studenttable.getModel().getValueAt(row, 2).toString());
						classname_text.setText(studenttable.getModel().getValueAt(row, 5).toString());
						classnumber_text.setText(studenttable.getModel().getValueAt(row, 6).toString());
						TakeCourse_Grade_text.setText(studenttable.getModel().getValueAt(row, 7).toString());
						grade_text.setText(studenttable.getModel().getValueAt(row, 8).toString());
						year_text.setText(studenttable.getModel().getValueAt(row, 3).toString());
						SemesterType_box.removeAllItems();
						if(studenttable.getModel().getValueAt(row, 4).toString().equals("上"))
	       				{
	    					SemesterType_box.addItem("上");
	    					SemesterType_box.addItem("下");
	    					SemesterType_box.addItem("--");
	       				}
	       				if(studenttable.getModel().getValueAt(row, 4).toString().equals("下"))
	       				{
	    					SemesterType_box.addItem("下");
	    					SemesterType_box.addItem("上");
	    					SemesterType_box.addItem("--");
	       				}
						
             }
		});
		
		determine.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	int flag = 0;
	        	int k;
	        	if(which==10)
	        	{
	        		int n = JOptionPane.showConfirmDialog(null, "確定匯出資料至Word檔?", "警告",JOptionPane.YES_NO_OPTION);
	        		if(n==0)
	        		{
	        			if(wordpath.equals(""))
	        			{
	        				JOptionPane.showMessageDialog(null, "範例Word檔路徑為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	        			}
	        			else
	        			{
	        				Path p = Paths.get(dir_path);
	        				if(!Files.exists(p))
	        				{
	        					try {
									Files.createDirectory(p);
								} catch (IOException e1) {
									e1.printStackTrace();
								}
	        				}
	        				
	        				try {
			            		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
			            		database_path = "jdbc:ucanaccess://"+datapath;
			            		connDB = DriverManager.getConnection(database_path);
								st=connDB.createStatement();
								for(int i=0;i<outputfile.size();i++)
									{
										k=0;
										do {
											if(repeat_output.size()==0)
											{
												flag = 1;
												break;
											}
											if(outputfile.get(i).toString().equals(repeat_output.get(k).toString()))
											{
												break;
											}
											if(k==repeat_output.size()-1)
											{
												flag = 1;
											}
											k++;
										}while(k<repeat_output.size());
										if(flag==1)
										{
											repeat_output.add(outputfile.get(i).toString());
											String output_sql = "Select * from 課程 where 課程代碼 = '"+outputfile.get(i).toString()+"'";
											rs = st.executeQuery(output_sql);
											while(rs.next())
											{
												printdata[0] = rs.getString("授課教師");
												printdata[1] = rs.getString("學年");
												printdata[2] = rs.getString("學期");
												printdata[3] = rs.getString("課程代碼");
												printdata[4] = rs.getString("課程名稱");
												printdata[5] = rs.getString("開課學制");
												printdata[6] = rs.getString("修習年級");
												printdata[7] = rs.getString("必選修");
												printdata[8] = rs.getString("平均");
												printdata[9] = rs.getString("修課人數");
												printdata[10] = rs.getString("不及格人數");
												printdata[11] = rs.getString("成績低於20人數");
												printdata[12] = rs.getString("及格率低於60");
												printdata[13] = rs.getString("平均不及格");
												pushword();
											}
											flag = 0;
										}
									}
								//JOptionPane.showMessageDialog(null, "已匯出成Word檔，共"+outputfile.size()+"筆");
								JOptionPane.showMessageDialog(null, "開始合併檔案，請稍等");
								data_sort();
								uniteDoc(list_combine,dir_path+"/" + "教學成效檢討回饋表.doc");
								JOptionPane.showMessageDialog(null, "合併完成");
							} catch (Exception e1) {
								e1.printStackTrace();
							}
	        			}
	        		}
	        	}
	        	else if(which==0)
	        	{
	        		if(excelpath.equals(""))
	        		{
	        			main.setSize(500,750);
	        			main.add(mainpage , BorderLayout.CENTER);
	        			mainpage.setVisible(true);
	        			datapage.setVisible(false);
	        			classpage.setVisible(false);
	        			studentpage.setVisible(false);
	        		}
	        		else if(!excelpath.equals("")&&datapath.equals(""))
	        		{
	        			JOptionPane.showMessageDialog(null, "資料庫路徑不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	        		}
	        		else
	        		{
	        				try {
	        					readXls(excelpath.toString());
	        					excelpath = "";
	        					excel_text.setText("");
	        				} catch (IOException e1) {
	        					
	        				} catch (SQLException e1) {
	        					e1.printStackTrace();
	        				}
	        	
	        				main.setSize(500,750);
	        				main.add(mainpage , BorderLayout.CENTER);
	        				mainpage.setVisible(true);
	        				datapage.setVisible(false);
	        				classpage.setVisible(false);
	    				studentpage.setVisible(false);
	        		}
	        	}
	        }
	    });
		
		returnmain.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	which = 4;
	        	main.setSize(500,750);
	        	main.add(mainpage , BorderLayout.CENTER);
	        	mainpage.setVisible(true);
	    		datapage.setVisible(false);
	    		classpage.setVisible(false);
	    		studentpage.setVisible(false);
	    		outputpage.setVisible(false);
	    		
	    		selectall.doClick();
	    		classname_text.setText("");
	    		classnumber_text.setText("");
	    		teachername_text.setText("");
	    		TakeCourse_Grade_text.setText("");
	    		studentname_text.setText("");
	    		studentnumber_text.setText("");
	    		year_text.setText("");
	    		grade_text.setText("");
	    		type_box.removeAllItems();
   				SemesterType_box.removeAllItems();
   				School_System_box.removeAllItems();
   				type_box.addItem("--");
   				type_box.addItem("必修");
				type_box.addItem("必選");
				type_box.addItem("選修");
				School_System_box.addItem("--");
				School_System_box.addItem("日間部");
				School_System_box.addItem("碩士班");
				School_System_box.addItem("碩士在職專班");
				SemesterType_box.addItem("--");
				SemesterType_box.addItem("上");
				SemesterType_box.addItem("下");
	        }
	    });
		
		count.addActionListener(new ActionListener() {
			 public void actionPerformed(ActionEvent e) {
				 int student=0,student_all=0,student_lower20=0,student_fail=0;
				 double average=0,average_all=0,fail_rate=0,pass_rate = 0;
				 Boolean b;
				 int flag=0;
				 int rounding_int = 0;
				 char rounding_str;
				 String average_string,average_all_string;
				 String fail_rate_string,average_fail_string,pass_rate_string;
				 try {
	    	            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
	    	            database_path = "jdbc:ucanaccess://"+datapath;
	    	            connDB = DriverManager.getConnection(database_path);
	    				st=connDB.createStatement();
				 for(int i=0;i<creatword.modelclass.getRowCount();i++)
             		{
             			b = Boolean.valueOf(creatword.modelclass.getValueAt(i, 0).toString());
             			if(b.booleanValue())
             			{
             				String sql = "select * from 學生資料 where 課程名稱 = '"+modelclass.getValueAt(i, 1).toString()+"' and 課程代碼='"+modelclass.getValueAt(i, 2).toString()+"' and 學年='"+modelclass.getValueAt(i,5).toString()+"' and 學期='"+modelclass.getValueAt(i, 6).toString()+"'";
             				rs = st.executeQuery(sql);
        					if(modelclass.getValueAt(i, 7).toString().equals("日間部"))
        					{
        						while(rs.next())
        						{
        							student_all++;
        							
        							average_all += Float.parseFloat(rs.getString("成績").trim());
        							if(Float.parseFloat(rs.getString("成績").trim())>20)
    								{
        								student++;
        								average += Float.parseFloat(rs.getString("成績").trim());
    								}
        							if(Float.parseFloat(rs.getString("成績").trim())<60)
        							{
        								student_fail++;
        							}
        							if(Float.parseFloat(rs.getString("成績").trim())<=20)
        							{
        								student_lower20++;
        							}
        						}
        						if(student_all!=0)
        						{
        							average_all /= student_all;
        						}
        						if(student!=0)
        						{
        							average /= student;	
        						}
        						fail_rate = (float)student_fail/student_all;
        						fail_rate*=100;
        						pass_rate = 100-fail_rate;
        						if(fail_rate>40)
        							fail_rate_string = "v";
        						else
        							fail_rate_string = "";
        						if(average<60)
        							average_fail_string = "v";
        						else
        							average_fail_string = "";
        						average = Math.round(average * 10.0)/10.0;
        						average_all = Math.round(average_all *10.0)/10.0;
        						fail_rate = Math.round(fail_rate*10.0)/10.0;
        						pass_rate = Math.round(pass_rate*10.0)/10.0;
    			        		String sql_find = "update 課程 set 平均='"+average+"',修課人數='"+student_all+"',不及格人數='"+student_fail+"',成績低於20人數='"+student_lower20+"',實際平均='"+average_all+"',及格率 = '"+pass_rate+"',及格率低於60 = '"+fail_rate_string+"',平均不及格 = '"+average_fail_string+"' where 課程代碼='"+modelclass.getValueAt(i, 2).toString()+"' and 學年 = '"+modelclass.getValueAt(i, 5).toString()+"' and 學期 = '"+modelclass.getValueAt(i, 6).toString()+"'";
    			        		ps = connDB.prepareStatement(sql_find);
    	    					ps.executeUpdate();
        						student=0;
        						student_all=0;
        						student_lower20=0;
        						student_fail=0;
        						average=0;
        						average_all=0;
        						fail_rate=0;
        						pass_rate=0;
        						flag=1;
								rounding_int=0;
        					}
        					else
        					{
        						while(rs.next())
        						{
        							student_all++;
        							average_all += Float.parseFloat(rs.getString("成績").trim());
        							if(Float.parseFloat(rs.getString("成績").trim())>20)
    								{
        								student++;
        								average += Float.parseFloat(rs.getString("成績").trim());
    								}
        							if(Float.parseFloat(rs.getString("成績").trim())<70)
        							{
        								student_fail++;
        							}
        							if(Float.parseFloat(rs.getString("成績").trim())<20)
        							{
        								student_lower20++;
        							}
        						}
        						if(student_all!=0)
        						{
        							average_all /= student_all;
        						}
        						if(student!=0)
        						{
        							average /= student;
        						}
        						fail_rate = (float)student_fail/student_all;
        						fail_rate*=100;
        						pass_rate = 100-fail_rate;
        						average = Math.round(average*10.0)/10.0;
        						average_all = Math.round(average_all*10.0)/10.0;
        						fail_rate = Math.round(fail_rate*10.0)/10.0;
        						pass_rate = Math.round(pass_rate*10.0)/10.0;
        						if(fail_rate>40)
        							fail_rate_string = "v";
        						else
        							fail_rate_string = "";
        						if(average<70)
        							average_fail_string = "v";
        						else
        							average_fail_string = "";
    			        		String sql_find = "update 課程 set 平均='"+average+"',修課人數='"+student+"',不及格人數='"+student_fail+"',成績低於20人數='"+student_lower20+"',實際平均='"+average_all+"',及格率 = '"+pass_rate+"',及格率低於60 = '"+fail_rate_string+"',平均不及格 = '"+average_fail_string+"' where 課程代碼='"+modelclass.getValueAt(i, 2).toString()+"'";
    			        		ps = connDB.prepareStatement(sql_find);
    	    					ps.executeUpdate();
        						student=0;
        						student_all=0;
        						student_lower20=0;
        						student_fail=0;
        						average=0;
        						average_all=0;
        						fail_rate=0;
        						pass_rate=0;
        						flag=1;
        					}
             			}
             		}
				 if(flag==0)
				 {
					 JOptionPane.showMessageDialog(null, "未選擇任何課程", "錯誤",JOptionPane.ERROR_MESSAGE); 
				 }
				 else
				 {
					 ps.close();
					 JOptionPane.showMessageDialog(null, "計算完成"); 
					 classname_text.setText("");
			    	 classnumber_text.setText("");
			    	 teachername_text.setText("");
			    	 TakeCourse_Grade_text.setText("");
			    	 year_text.setText("");
			    	 type_box.removeAllItems();
		   			 SemesterType_box.removeAllItems();
		   			 School_System_box.removeAllItems();
		   			 type_box.addItem("--");
		   			 type_box.addItem("必修");
					 type_box.addItem("必選");
					 type_box.addItem("選修");
					 School_System_box.addItem("--");
					 School_System_box.addItem("日間部");
					 School_System_box.addItem("碩士班");
					 School_System_box.addItem("碩士在職專班");
					 SemesterType_box.addItem("--");
					 SemesterType_box.addItem("上");
					 SemesterType_box.addItem("下");
					 selectall.doClick();
					 find.doClick();
				 }
				 }catch (SQLException e1) {
 					e1.printStackTrace();
 				} catch (ClassNotFoundException e1) {
 					e1.printStackTrace();
 				}	
			 }
		});
		
		delete.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	Boolean b;
	        	try {
    	            Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
    	            database_path = "jdbc:ucanaccess://"+datapath;
    	            connDB = DriverManager.getConnection(database_path);
    				st=connDB.createStatement();
    					if(which==1)
    			        {
    			        	int n = JOptionPane.showConfirmDialog(null, "是否刪除資料? 刪除後將無法復原", "警告",JOptionPane.YES_NO_OPTION);
    		        		if(n==0)
    			        	{
    		        			for(int i=0;i<creatword.modelclass.getRowCount();i++)
    		        			{
    		        				b = Boolean.valueOf(creatword.modelclass.getValueAt(i, 0).toString());
    		        				if(b.booleanValue())
    		        				{
    		        					String sql = "delete from 課程 where 課程名稱 = '"+modelclass.getValueAt(i, 1).toString()+"' and 課程代碼='"+modelclass.getValueAt(i, 2).toString()+"' and 學年 = '"+modelclass.getValueAt(i, 5).toString()+"' and 學期 = '"+modelclass.getValueAt(i, 6).toString()+"'";
    		        					ps = connDB.prepareStatement(sql);
    		        					ps.executeUpdate();
    		        				}
    		        			}
    		        				classname_text.setText("");
    		        				classnumber_text.setText("");
    		        				teachername_text.setText("");
    		        				TakeCourse_Grade_text.setText("");
    		        				year_text.setText("");
    		        				type_box.removeAllItems();
    		        				SemesterType_box.removeAllItems();
    		        				School_System_box.removeAllItems();
    		        				type_box.addItem("--");
    		        				type_box.addItem("必修");
    		        				type_box.addItem("必選");
    		        				type_box.addItem("選修");
    		        				School_System_box.addItem("--");
	    							School_System_box.addItem("日間部");
	    							School_System_box.addItem("碩士班");
	    							School_System_box.addItem("碩士在職專班");
	    							SemesterType_box.addItem("--");
	    							SemesterType_box.addItem("上");
	    							SemesterType_box.addItem("下");
	    							find.doClick();
    		        				ps.close();
    		        				if(classtable.getModel().getRowCount()==0)
    		    	        		{
    		    	        			selectall.setText("全選");
    		    	        		}
    		        				JOptionPane.showMessageDialog(null, "刪除成功");
    			        	}
    			        }
    					if(which==2)
    		        	{
    						int n = JOptionPane.showConfirmDialog(null, "是否刪除資料? 刪除後將無法復原", "警告",JOptionPane.YES_NO_OPTION);
    		        		if(n==0)
    			        	{
    		        			for(int i=0;i<creatword.modelstudent.getRowCount();i++)
    		        			{
    		        				b = Boolean.valueOf(creatword.modelstudent.getValueAt(i, 0).toString());
    		        				if(b.booleanValue())
    		        				{
    		        					String sql = "delete from 學生資料 where 姓名 = '"+modelstudent.getValueAt(i, 1).toString()+"' and 學號='"+modelstudent.getValueAt(i, 2).toString()+"' and 課程代碼 = '"+modelstudent.getValueAt(i, 6).toString()+"' and 學年 = '"+modelstudent.getValueAt(i, 3).toString()+"' and 學期 = '"+modelstudent.getValueAt(i, 4).toString()+"'";
    		        					ps = connDB.prepareStatement(sql);
    		        					ps.executeUpdate();
    		        				}
    		        			}
    		        				classname_text.setText("");
    		        				classnumber_text.setText("");
    		        				TakeCourse_Grade_text.setText("");
    		        				studentname_text.setText("");
    		        				studentnumber_text.setText("");
    		        				grade_text.setText("");
    		        				year_text.setText("");
    		        				SemesterType_box.removeAllItems();
    		        				SemesterType_box.addItem("--");
    		        				SemesterType_box.addItem("上");
    		        				SemesterType_box.addItem("下");
	    							find.doClick();
    		        				ps.close();
    		        				if(studenttable.getModel().getRowCount()==0)
    		    	        		{
    		    	        			selectall.setText("全選");
    		    	        		}
    		        				JOptionPane.showMessageDialog(null, "刪除成功"); 
    			        	}
    		        	}
    				} catch (SQLException e1) {
    					// TODO Auto-generated catch block
    					e1.printStackTrace();
    				} catch (ClassNotFoundException e1) {
    					e1.printStackTrace();
    				}
	        	if(which==10)
	        	{
	        		outputtable.addMouseListener(new MouseAdapter() {
	        			public void mouseClicked(MouseEvent e) {
	        				row = outputtable.getSelectedRow();
	        			}
	        		});
	        		String number = outputtable.getModel().getValueAt(row, 1).toString();
    				for(int i=0;i<outputfile.size();i++)
    				{
    					if(number.equals(outputfile.get(i)))
    						{
    							outputfile.remove(i);
    							JOptionPane.showMessageDialog(null, "刪除成功"); 
    						}
    				}
	        		find.doClick();
	        	}
	        }
	    });
		
		add.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	if(which==1)
	        	{
	        		if(classnumber_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "課程名稱不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(classname_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "課程代碼不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(teachername_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "教師名稱不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(TakeCourse_Grade_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "修習年級不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(year_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "學年不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(type_box.getSelectedItem().toString().equals("--"))
	            	{
	        			JOptionPane.showMessageDialog(null, "必選修不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(School_System_box.getSelectedItem().toString().equals("--"))
	            	{
	        			JOptionPane.showMessageDialog(null, "開課學制不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(SemesterType_box.getSelectedItem().toString().equals("--"))
	            	{
	        			JOptionPane.showMessageDialog(null, "學期不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else
	        		{
	        			int n = JOptionPane.showConfirmDialog(null, "確定新增資料?", "警告",JOptionPane.YES_NO_OPTION);
	        			if(n==0)
		        		{
	        				try {
	    	            		creatword.modelclass.setRowCount(0);
	    	            		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
	    	            		database_path = "jdbc:ucanaccess://"+datapath;
	    	            		connDB = DriverManager.getConnection(database_path);
	    						st=connDB.createStatement();
	    						String sql = "INSERT INTO 課程 (課程名稱,課程代碼,授課教師,修習年級,必選修,學年,學期,開課學制)  VALUES('"+classname_text.getText().toString()+"','"+classnumber_text.getText().toString()+"','"+teachername_text.getText().toString()+"','"+TakeCourse_Grade_text.getText().toString()+"','"+type_box.getSelectedItem().toString()+"','"+year_text.getText().toString()+"','"+SemesterType_box.getSelectedItem().toString()+"','"+School_System_box.getSelectedItem().toString()+"')";
	    						ps = connDB.prepareStatement(sql);
	    						ps.executeUpdate();
	    						ps.close();
	    						classname_text.setText("");
	    			    		classnumber_text.setText("");
	    			    		teachername_text.setText("");
	    			    		TakeCourse_Grade_text.setText("");
	    			    		year_text.setText("");
	    			    		type_box.removeAllItems();
	    		   				SemesterType_box.removeAllItems();
	    		   				School_System_box.removeAllItems();
	    		   				type_box.addItem("--");
	    		   				type_box.addItem("必修");
	    						type_box.addItem("必選");
	    						type_box.addItem("選修");
	    						School_System_box.addItem("--");
	    						School_System_box.addItem("日間部");
	    						School_System_box.addItem("碩士班");
	    						School_System_box.addItem("碩士在職專班");
	    						SemesterType_box.addItem("--");
	    						SemesterType_box.addItem("上");
	    						SemesterType_box.addItem("下");
	    						find.doClick();
	    						JOptionPane.showMessageDialog(null, "新增成功"); 
	    					} catch (SQLException e1) {
	    						JOptionPane.showMessageDialog(null, "新增有誤,請檢查後重新新增", "錯誤",JOptionPane.ERROR_MESSAGE);
	    						e1.printStackTrace();
	    					} catch (ClassNotFoundException e1) {
	    						e1.printStackTrace();
	    					}
	        				
		        		}
	        		}
	        	}
	        	if(which==2)
	        	{
	        		if(classnumber_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "課程名稱不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(classname_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "課程代碼不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(studentname_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "學生姓名不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(TakeCourse_Grade_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "修習年級不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(studentnumber_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "學號不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(grade_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "成績不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(year_text.getText().equals(""))
	            	{
	        			JOptionPane.showMessageDialog(null, "學年不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else if(SemesterType_box.getSelectedItem().toString().equals("--"))
	            	{
	        			JOptionPane.showMessageDialog(null, "學期不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
	            	}
	        		else
	        		{
	        			int n = JOptionPane.showConfirmDialog(null, "確定新增資料?", "警告",JOptionPane.YES_NO_OPTION);
	        			if(n==0)
		        		{
	        				try {
	    	            		creatword.modelclass.setRowCount(0);
	    	            		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
	    	            		database_path = "jdbc:ucanaccess://"+datapath;
	    	            		connDB = DriverManager.getConnection(database_path);
	    						st=connDB.createStatement();
	    						String sql = "INSERT INTO 學生資料 (學年,學期,學號,姓名,成績,課程名稱,課程代碼,修課年級)  VALUES('"+year_text.getText().toString()+"','"+SemesterType_box.getSelectedItem().toString()+"','"+studentnumber_text.getText().toString()+"','"+studentname_text.getText().toString()+"','"+grade_text.getText().toString()+"','"+classname_text.getText().toString()+"','"+classnumber_text.getText().toString()+"','"+TakeCourse_Grade_text.getText()+"')";
	    						ps = connDB.prepareStatement(sql);
	    						ps.executeUpdate();
	    						ps.close();
	    						classname_text.setText("");
	    			    		classnumber_text.setText("");
	    			    		TakeCourse_Grade_text.setText("");
	    			    		studentname_text.setText("");
	    			    		studentnumber_text.setText("");
	    			    		grade_text.setText("");
	    			    		year_text.setText("");
	    			    		SemesterType_box.removeAllItems();
	    			    		SemesterType_box.addItem("--");
	    						SemesterType_box.addItem("上");
	    						SemesterType_box.addItem("下");
	    						find.doClick();
	    						JOptionPane.showMessageDialog(null, "新增成功"); 
	    					} catch (SQLException e1) {
	    						e1.printStackTrace();
	    					} catch (ClassNotFoundException e1) {
	    						e1.printStackTrace();
	    					}
	        				
		        		}
	        		}
	        	}
	        }
	    });
		
		modify.addActionListener(new ActionListener() {
	        public void actionPerformed(ActionEvent e) {
	        	try {
					Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
					database_path = "jdbc:ucanaccess://"+datapath;
	        		connDB = DriverManager.getConnection(database_path);
					st=connDB.createStatement();
					String modifyclass;
					String modifystudent;
					if(which==1)
		        	{
		        		row = classtable.getSelectedRow();
		        		if(classnumber_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "課程名稱不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(classname_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "課程代碼不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(teachername_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "教師名稱不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(TakeCourse_Grade_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "修習年級不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(year_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "學年不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(type_box.getSelectedItem().toString().equals("--"))
		            	{
		        			JOptionPane.showMessageDialog(null, "必選修不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(School_System_box.getSelectedItem().toString().equals("--"))
		            	{
		        			JOptionPane.showMessageDialog(null, "開課學制不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(SemesterType_box.getSelectedItem().toString().equals("--"))
		            	{
		        			JOptionPane.showMessageDialog(null, "學期不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else
		        		{
		        			int n = JOptionPane.showConfirmDialog(null, "確定修改資料?", "警告",JOptionPane.YES_NO_OPTION);
		        			if(n==0)
			        		{
			        			modifyclass = "update 課程 set 課程名稱='"+classname_text.getText().toString()+"', 課程代碼='"+classnumber_text.getText().toString()+"', 授課教師='"+teachername_text.getText().toString()+"', 修習年級='"+TakeCourse_Grade_text.getText().toString()+"', 必選修='"+type_box.getSelectedItem().toString()+"', 學年='"+year_text.getText().toString()+"',學期='"+SemesterType_box.getSelectedItem().toString()+"', 開課學制='"+School_System_box.getSelectedItem().toString()+"' where 課程代碼='"+classtable.getModel().getValueAt(row, 2).toString()+"' and 學年 = '"+classtable.getModel().getValueAt(row, 5)+"' and 學期 = '"+classtable.getModel().getValueAt(row, 6)+"'";
			        			ps = connDB.prepareStatement(modifyclass);
								ps.executeUpdate();
								modifystudent = "update 學生資料 set 課程名稱='"+classname_text.getText().toString()+"', 課程代碼='"+classnumber_text.getText().toString()+"', 學年='"+year_text.getText().toString()+"', 學期='"+SemesterType_box.getSelectedItem().toString()+"' where 課程代碼='"+classtable.getModel().getValueAt(row, 2).toString()+"' and 學年 = '"+classtable.getModel().getValueAt(row, 5)+"' and 學期 = '"+classtable.getModel().getValueAt(row, 6)+"'";
								ps = connDB.prepareStatement(modifystudent);
								ps.executeUpdate();
								ps.close();
								classname_text.setText("");
	    			    		classnumber_text.setText("");
	    			    		teachername_text.setText("");
	    			    		TakeCourse_Grade_text.setText("");
	    			    		year_text.setText("");
	    			    		type_box.removeAllItems();
	    		   				SemesterType_box.removeAllItems();
	    		   				School_System_box.removeAllItems();
	    		   				type_box.addItem("--");
	    		   				type_box.addItem("必修");
	    						type_box.addItem("必選");
	    						type_box.addItem("選修");
	    						School_System_box.addItem("--");
	    						School_System_box.addItem("日間部");
	    						School_System_box.addItem("碩士班");
	    						School_System_box.addItem("碩士在職專班");
	    						SemesterType_box.addItem("--");
	    						SemesterType_box.addItem("上");
	    						SemesterType_box.addItem("下");
	    						find.doClick();
	    						JOptionPane.showMessageDialog(null, "修改成功");  
			        		}
		        		}
		        	}
		        	if(which==2)
		        	{
		        		row = studenttable.getSelectedRow();
		        		if(classnumber_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "課程名稱不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(classname_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "課程代碼不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(studentname_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "學生姓名不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(TakeCourse_Grade_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "修習年級不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(studentnumber_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "學號不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(grade_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "成績不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(year_text.getText().equals(""))
		            	{
		        			JOptionPane.showMessageDialog(null, "學年不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else if(SemesterType_box.getSelectedItem().toString().equals("--"))
		            	{
		        			JOptionPane.showMessageDialog(null, "學期不可為空", "錯誤",JOptionPane.ERROR_MESSAGE); 
		            	}
		        		else
		        		{
		        			int n = JOptionPane.showConfirmDialog(null, "確定修改資料?", "警告",JOptionPane.YES_NO_OPTION);
		        			if(n==0)
			        		{
		        				modifystudent = "update 學生資料 set 學年='"+year_text.getText().toString()+"', 學期='"+SemesterType_box.getSelectedItem().toString()+"', 學號='"+studentnumber_text.getText().toString()+"', 姓名='"+studentname_text.getText().toString()+"', 成績='"+grade_text.getText().toString()+"', 課程名稱='"+classname_text.getText().toString()+"', 課程代碼='"+classnumber_text.getText().toString()+"',修課年級='"+TakeCourse_Grade_text.getText().toString()+"' where 學號='"+studenttable.getModel().getValueAt(row, 2).toString()+"' and 課程代碼='"+studenttable.getModel().getValueAt(row, 6).toString()+"' and 學年 = '"+studenttable.getModel().getValueAt(row, 3)+"' and 學期 = '"+studenttable.getModel().getValueAt(row, 4)+"'";
		        				ps = connDB.prepareStatement(modifystudent);
		    					ps.executeUpdate();
		    					classname_text.setText("");
		    			    	classnumber_text.setText("");
		    			    	TakeCourse_Grade_text.setText("");
		    			    	studentname_text.setText("");
		    			    	studentnumber_text.setText("");
		    			    	grade_text.setText("");
		    			    	year_text.setText("");
		    			    	SemesterType_box.removeAllItems();
		    			    	SemesterType_box.addItem("--");
		    					SemesterType_box.addItem("上");
		    					SemesterType_box.addItem("下");
		    					find.doClick();
		    					JOptionPane.showMessageDialog(null, "修改成功"); 
			        		}
		        		}
		        	}
				} catch (ClassNotFoundException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (SQLException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
	        }
	    });
		
		inputdatabase.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	 creatword.filechooser = new JFileChooser(path);
	        	 creatword.filechooser.setMultiSelectionEnabled(true);
	        	 creatword.filechooser.setDialogTitle("請選擇資料庫檔案");
	        	 creatword.filechooser.setFileFilter(access);
	        	 files = null;
	             int returnVal = creatword.filechooser.showOpenDialog(null);
	             if (returnVal == 0) {
	            	 creatword.filePath = creatword.filechooser.getSelectedFile().getAbsolutePath();
	            	 creatword.files = creatword.filechooser.getSelectedFile();
	            	 datapath = creatword.files.getAbsolutePath().toString();
	            	 database_text.setText(datapath);
	             }
	        }
	    });
		
		inputexcel.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	 creatword.filechooser = new JFileChooser(path);
	        	 creatword.filechooser.setMultiSelectionEnabled(true);
	        	 creatword.filechooser.setDialogTitle("請選擇Excel檔案");
	        	 creatword.filechooser.setFileFilter(xlsx);
	        	 files = null;
	             int returnVal = creatword.filechooser.showOpenDialog(null);
	             if (returnVal == 0) {
	            	 creatword.filePath = creatword.filechooser.getSelectedFile().getAbsolutePath();
	            	 creatword.files = creatword.filechooser.getSelectedFile();
 	            	 excelpath = creatword.files.getAbsolutePath().toString();
	            	 excel_text.setText(excelpath);
	             }
	        }
	    });
		
		inputword.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	 creatword.filechooser = new JFileChooser(path);
	        	 creatword.filechooser.setMultiSelectionEnabled(true);
	        	 creatword.filechooser.setDialogTitle("請選擇Wor範例檔案");
	        	 creatword.filechooser.setFileFilter(docx);
	        	 files = null;
	             int returnVal = creatword.filechooser.showOpenDialog(null);
	             if (returnVal == 0) {
	            	 creatword.filePath = creatword.filechooser.getSelectedFile().getAbsolutePath();
	            	 creatword.files = creatword.filechooser.getSelectedFile();
	            	 wordpath= creatword.files.getAbsolutePath().toString();
	            	 word_text.setText(wordpath);
	             }
	        }
	    });
		
		output.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        			Boolean b;
	        			int count_error = 0;
    			        	int n = JOptionPane.showConfirmDialog(null, "是否加入匯出檔案? ", "警告",JOptionPane.YES_NO_OPTION);
    		        		if(n==0)
    			        	{
    		        			for(int i=0;i<creatword.modelclass.getRowCount();i++)
    		        			{
    		        				b = Boolean.valueOf(creatword.modelclass.getValueAt(i, 0).toString());
    		        				if(b.booleanValue())
    		        				{
    		        					for(int j=0;j<outputfile.size();j++)
    		        					{
    		        						if(creatword.modelclass.getValueAt(i, 2).toString().equals(outputfile.get(j).toString())&&creatword.modelclass.getValueAt(i, 5).toString().equals(outputfile_year.get(j).toString())&&creatword.modelclass.getValueAt(i, 6).toString().equals(outputfile_semester.get(j).toString()))
    		        						{
    		        							JOptionPane.showMessageDialog(null, "重複加入檔案 學年:"+creatword.modelclass.getValueAt(i, 5).toString()+" 學期:"+creatword.modelclass.getValueAt(i, 6).toString()+" 課程代碼:"+creatword.modelclass.getValueAt(i, 2).toString(), "警告",JOptionPane.ERROR_MESSAGE); 
    		        							break;
    		        						}
    		        						else {
    		        							count_error+=1;
    		        						}
    		        					}
    		        					if(count_error==outputfile.size())
    		        					{
    		        						outputfile.add(creatword.modelclass.getValueAt(i, 2).toString());//課程代號
    		        						outputfile_year.add(creatword.modelclass.getValueAt(i, 5).toString());//學年
    		        						outputfile_semester.add(creatword.modelclass.getValueAt(i, 6).toString());//學期
    		        						if(outputfile_teacher.size()==0) {
    		        							outputfile_teacher.add(creatword.modelclass.getValueAt(i, 9).toString());
    		        						}else {
    		        							int nrepeat = 0;
    		        							for(int teachers=0;teachers<outputfile_teacher.size();teachers++)
        		        						{
    		        								if(!creatword.modelclass.getValueAt(i, 9).toString().equals(outputfile_teacher.get(teachers).toString())) {
    		        									nrepeat++;
    		        								}
        		        						}
    		        							if(nrepeat==outputfile_teacher.size()) {
    		        								outputfile_teacher.add(creatword.modelclass.getValueAt(i, 9).toString());
    		        							}
    		        						}
    		        					}
    		        					count_error = 0;
    		        				}
    		        			}
    		        			JOptionPane.showMessageDialog(null, "已更新匯出列表");
    			        	}
	        }
	    });
		
		find.addActionListener(new ActionListener() {
	        @Override
	        public void actionPerformed(ActionEvent e) {
	        	if(datapath.equals(""))
	        	{
	        		JOptionPane.showMessageDialog(null, "尚未選擇資料庫路徑");
	        		returnmain.doClick();
	        	}
	        	else
	        	{
	        	creatword.modelclass.setRowCount(0);
	        	creatword.modelstudent.setRowCount(0);
	        	String findclass = "select * from 課程 " ;
	        	String findstudent = "select * from 學生資料 ";
	        	String where = "where true";
	        	String add = "";
	        	String classname,classnumber,teachername,TakeCourse_Grade,year,type,School_System,SemesterType;
	        	String studentname,studentnumber,grade,average;
	        	String total_student,fail_grade,lower_grade;
	             if(which==1)//課程
	             { 
	            	if(!classnumber_text.getText().equals(""))
	            	{
	            		where = "where true";
	            		add+=" AND 課程代碼 LIKE  '%"+classnumber_text.getText().toString()+"%'";
	            	}
	            	if(!classname_text.getText().equals(""))
	            	{
	            		where = "where true";
	            		add+=" AND 課程名稱 LIKE   '%"+classname_text.getText().toString()+"%'";
	            	}
	            	if(!teachername_text.getText().equals(""))
	            	{
	            		where = "where true";
	            		add+=" AND 授課教師 LIKE   '%"+teachername_text.getText().toString()+"%'";
	            	}
	            	if(!TakeCourse_Grade_text.getText().equals(""))
	            	{
	            		where = "where true";
	            		add+=" AND 修習年級 LIKE   '%"+TakeCourse_Grade_text.getText().toString()+"%'";
	            	}
	            	if(!year_text.getText().equals(""))
	            	{
	            		where = "where true";
	            		add+=" AND 學年 LIKE  '%"+year_text.getText().toString()+"%'";
	            	}
	            	if(!type_box.getSelectedItem().toString().equals("--"))
	            	{
	            		where = "where true";
	            		add+=" AND 必選修 LIKE  '%"+type_box.getSelectedItem().toString()+"%'";
	            	}
	            	if(!School_System_box.getSelectedItem().toString().equals("--"))
	            	{
	            		where = "where true";
	            		add+=" AND 開課學制 LIKE  '%"+School_System_box.getSelectedItem().toString()+"%'";
	            	}
	            	if(!SemesterType_box.getSelectedItem().toString().equals("--"))
	            	{
	            		where = "where true";
	            		add+=" AND 學期 LIKE  '%"+SemesterType_box.getSelectedItem().toString()+"%'";
	            	}
	            	findclass = findclass+where+add;
	            	try {
	            		creatword.modelclass.setRowCount(0);
	            		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
	            		database_path = "jdbc:ucanaccess://"+datapath;
	            		connDB = DriverManager.getConnection(database_path);
						st=connDB.createStatement();
						rs = st.executeQuery(findclass);
						while(rs.next())
						{
							classname = rs.getString("課程名稱");
							classnumber = rs.getString("課程代碼");
							teachername = rs.getString("授課教師");
							TakeCourse_Grade = rs.getString("修習年級");
							year = rs.getString("學年");
							type = rs.getString("必選修");
							School_System = rs.getString("開課學制");
							SemesterType = rs.getString("學期");
							average = rs.getString("平均");
							total_student = rs.getString("修課人數");
							fail_grade = rs.getString("不及格人數");
							lower_grade = rs.getString("成績低於20人數");
							Object[] row = { new JCheckBox(), classname,classnumber,type,average,year,SemesterType,School_System,TakeCourse_Grade,teachername,total_student,fail_grade,lower_grade};
							creatword.modelclass.addRow(row);
							
						}
					} catch (SQLException e1) {
						e1.printStackTrace();
					} catch (ClassNotFoundException e1) {
						e1.printStackTrace();
					}
					
	             }
	             else if(which==2)//學生資料
	             {
	            	 if(!classnumber_text.getText().equals(""))
		            	{
		            		where = "where true";
		            		add+=" AND 課程代碼 LIKE '%"+classnumber_text.getText().toString()+"%'";
		            	}
		            	if(!classname_text.getText().equals(""))
		            	{
		            		where = "where true";
		            		add+=" AND 課程名稱 LIKE  '%"+classname_text.getText().toString()+"%'";
		            	}
		            	if(!studentnumber_text.getText().equals(""))
		            	{
		            		where = "where true";
		            		add+=" AND 學號 LIKE  '%"+studentnumber_text.getText().toString()+"%'";
		            	}
		            	if(!studentname_text.getText().equals(""))
		            	{
		            		where = "where true";
		            		add+=" AND 姓名 LIKE '%"+studentname_text.getText().toString()+"%'";
		            	}
		            	if(!TakeCourse_Grade_text.getText().equals(""))
		            	{
		            		where = "where true";
		            		add+=" AND 修課年級 LIKE  '%"+TakeCourse_Grade_text.getText().toString()+"%'";
		            	}
		            	if(!grade_text.getText().equals(""))
		            	{
		            		where = "where true";
		            		add+=" AND 成績 LIKE  '%"+grade_text.getText().toString()+"%'";
		            	}
		            	if(!year_text.getText().equals(""))
		            	{
		            		where = "where true";
		            		add+=" AND 學年 LIKE  '%"+year_text.getText().toString()+"%'";
		            	}
		            	if(!SemesterType_box.getSelectedItem().toString().equals("--"))
		            	{
		            		where = "where true";
		            		add+=" AND 學期 LIKE  '%"+SemesterType_box.getSelectedItem().toString()+"%'";
		            	}
		            	findstudent = findstudent+where+add;
		            	try {
		            		creatword.modelclass.setRowCount(0);
		            		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
		            		database_path = "jdbc:ucanaccess://"+datapath;
		            		connDB = DriverManager.getConnection(database_path);
							st=connDB.createStatement();
							rs = st.executeQuery(findstudent);
							while(rs.next())
							{
								classname = rs.getString("課程名稱");
								classnumber = rs.getString("課程代碼");
								studentname = rs.getString("姓名");
								studentnumber = rs.getString("學號");
								TakeCourse_Grade = rs.getString("修課年級");
								grade = rs.getString("成績");
								year = rs.getString("學年");
								SemesterType = rs.getString("學期");
								Object[] row = { new JCheckBox(), studentname,studentnumber,year,SemesterType,classname,classnumber,TakeCourse_Grade,grade};
								creatword.modelstudent.addRow(row);
							}
							
						} catch (SQLException e1) {
							e1.printStackTrace();
						} catch (ClassNotFoundException e1) {
							e1.printStackTrace();
						}
	             }
	             else if(which==10)
	             {
		            	try {
		            		creatword.modeloutput.setRowCount(0);
		            		Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
		            		database_path = "jdbc:ucanaccess://"+datapath;
		            		connDB = DriverManager.getConnection(database_path);
							st=connDB.createStatement();
							if(outputfile.size()==0)
							{
								JOptionPane.showMessageDialog(null, "無任何匯出資料", "錯誤",JOptionPane.ERROR_MESSAGE);
								returnmain.doClick();
							}
							else
								{
								for(int i=0;i<outputfile.size();i++)
									{
										String output_sql = "Select * from 課程 where 課程代碼 = '"+outputfile.get(i).toString()+"' and 學年 = '"+outputfile_year.get(i).toString()+"' and 學期 = '"+outputfile_semester.get(i).toString()+"'";
										rs = st.executeQuery(output_sql);
										rs.next();
										classname = rs.getString("課程名稱");
										classnumber = rs.getString("課程代碼");
										teachername = rs.getString("授課教師");
										TakeCourse_Grade = rs.getString("修習年級");
										year = rs.getString("學年");
										type = rs.getString("必選修");
										School_System = rs.getString("開課學制");
										SemesterType = rs.getString("學期");
										average = rs.getString("平均");
										total_student = rs.getString("修課人數");
										fail_grade = rs.getString("不及格人數");
										lower_grade = rs.getString("成績低於20人數");
										Object[] row = {classname,classnumber,type,average,year,SemesterType,School_System,TakeCourse_Grade,teachername,total_student,fail_grade,lower_grade};
										creatword.modeloutput.addRow(row);
									}
								}
						} catch (SQLException e1) {
							e1.printStackTrace();
						} catch (ClassNotFoundException e1) {
							e1.printStackTrace();
						}
	             }
	        }
	        }
	    });
		
		clearfile.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent e) {
				try {
					int n = JOptionPane.showConfirmDialog(null, "是否清空資料表,如果清空將無法還原. 確定清除?", "警告",JOptionPane.YES_NO_OPTION);
					String sql;
					if(n==0)
					{
						Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
						database_path = "jdbc:ucanaccess://"+datapath;
	            		connDB = DriverManager.getConnection(database_path);
	            		//Access 不支援 TRUNCATE TABLE 所以使用 delete from
						sql = "delete from 課程";
						ps = connDB.prepareStatement(sql);
						ps.executeUpdate();
						sql = "delete from 學生資料";
						ps = connDB.prepareStatement(sql);
						ps.executeUpdate();
						ps.close();
						connDB.close();
						JOptionPane.showMessageDialog(null, "已清空"); 
					}
				} catch (SQLException e1) {
					e1.printStackTrace();
				} catch (ClassNotFoundException e1) {
					e1.printStackTrace();
				}
			}
		});
		
		selectall.addActionListener(new ActionListener() {
			private int sum;
	        public void actionPerformed(ActionEvent e) {
	        	if(which==1)//課程
	            {
	        		if (this.sum == 0) {
	                    for (int j = 0; j < classtable.getModel().getRowCount(); j++) {
	                    	classtable.getModel().setValueAt(Boolean.valueOf(true), j, 0);
	                    	selectall.setText("取消全選");
	                    } 
	                  } else {
	                    for (int j = 0; j < classtable.getModel().getRowCount(); j++) {
	                    	classtable.getModel().setValueAt(Boolean.valueOf(false), j, 0);
	                      this.sum = 0;
	                      selectall.setText("全選");
	                    } 
	                  } 
	                  for (int i = 0; i < classtable.getModel().getRowCount(); i++) {
	                    if (((Boolean)classtable.getModel().getValueAt(i, 0)).booleanValue())
	                      this.sum++; 
	                  } 
	            }
	            else if(which==2)//學生資料count
	            {
	            	if (this.sum == 0) {
	                    for (int j = 0; j < studenttable.getModel().getRowCount(); j++) {
	                    	studenttable.getModel().setValueAt(Boolean.valueOf(true), j, 0);
	                    	selectall.setText("取消全選");
	                    } 
	                  } else {
	                    for (int j = 0; j < studenttable.getModel().getRowCount(); j++) {
	                    	studenttable.getModel().setValueAt(Boolean.valueOf(false), j, 0);
	                      this.sum = 0;
	                      selectall.setText("全選");
	                    } 
	                  } 
	                  for (int i = 0; i < studenttable.getModel().getRowCount(); i++) {
	                    if (((Boolean)studenttable.getModel().getValueAt(i, 0)).booleanValue())
	                      this.sum++; 
	                  } 
	            }
	            else if(which==4)
	            {
	            	for (int j = 0; j < studenttable.getModel().getRowCount(); j++) {
                    	studenttable.getModel().setValueAt(Boolean.valueOf(false), j, 0);
                      this.sum = 0;
                      selectall.setText("全選");
                    } 
	            	for (int j = 0; j < classtable.getModel().getRowCount(); j++) {
	                    classtable.getModel().setValueAt(Boolean.valueOf(false), j, 0);
	                    this.sum = 0;
	                    selectall.setText("全選");
	                }
	            }
	        }
	    });
		
		main.setResizable(false);
		main.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		main.setVisible(true);
	}
	
	public static void database() throws SQLException
	{
		// 選課代碼 學年 學期為主鍵
		try
		{
			int j=0;
			int i=0;
			int k=0;
			int flag = 0;
			int count = 0;
			ArrayList varify_class = new ArrayList();
			ArrayList varify_year = new ArrayList();
			ArrayList varify_semester = new ArrayList();
			ArrayList varify_number = new ArrayList();
			ArrayList total_class = new ArrayList();
			ArrayList datafromexcel = new ArrayList();
			Class.forName("net.ucanaccess.jdbc.UcanaccessDriver");
			database_path = "jdbc:ucanaccess://"+datapath;
			connDB = DriverManager.getConnection(database_path);
			Statement st=connDB.createStatement();
				for(k=0;k<excelrow-2;k++)
				{
					for(i=0;i<excelcolumn;i++)
					{
						if(i==0)
						{
							//學年位置 去除小數點
								if(exceldata.get(j).toString().length()>3)
							{
								exceldata.set(j, exceldata.get(j).toString().substring(0,3));
							}
						}
						if(i==2)
						{
							//選課代碼位置 去除小數點
							if(exceldata.get(j).toString().length()>4)
							{
								exceldata.set(j, exceldata.get(j).toString().substring(0,4));
							}
						}
						if(i==7)
						{
							//學分數位置 去除小數點
							if(exceldata.get(j).toString().length()>1)
							{
								exceldata.set(j, exceldata.get(j).toString().substring(0,1));
							}
						}
						j++;
					}
				}
				//找出課程
				for(k=0;k<exceldata.size();k++)
				{
					i=0;
					varify_class.add(exceldata.get(k));
					if((k+1)%excelcolumn==0)
					{
						do
						{
							if(varify_number.size()==0)
							{
								flag = 1;
								varify_class.removeAll(varify_class);
								break;
							}
							else
							{
								if(varify_class.get(2).toString().equals(varify_number.get(i).toString()))
									{
									if(varify_class.get(0).toString().equals(varify_year.get(i).toString()))
									{
										if(varify_class.get(1).toString().equals(varify_semester.get(i).toString()))
										{
											varify_class.removeAll(varify_class);
											break;
										}
									}
									}
								if(i==varify_number.size()-1)
								{
									flag = 1;
								}
							}
							i++;
						}
						while(i<varify_number.size());
						if(flag==1)
						{
							for(i=(k-excelcolumn)+1;i<k;i++)
							{
								// 0 學年 1 學期 2 選課代碼 3 開課學制 4 授課教師 5 修習年級 6 必選修 7 學分數 11 課程名稱
								switch (count){
									case 0:
										varify_year.add(exceldata.get(i));
										total_class.add(exceldata.get(i));
										break;
									case 1:
										varify_semester.add(exceldata.get(i));
										total_class.add(exceldata.get(i));
										break;
									case 2:
										varify_number.add(exceldata.get(i));
										total_class.add(exceldata.get(i));
										break;
									case 3:
										total_class.add(exceldata.get(i));
										break;
									case 4:
										total_class.add(exceldata.get(i));
										break;
									case 5:
										total_class.add(exceldata.get(i));
										break;
									case 6:
										total_class.add(exceldata.get(i));
										break;
									case 7:
										total_class.add(exceldata.get(i));
										break;
									case 11:
										total_class.add(exceldata.get(i));
										break;
								}
								count++;
							}
							count = 0;
							flag = 0;
						}
						varify_class.removeAll(varify_class);
					}
				}
				/*Object[] obj = new Object[excelrow];
				String sql = "INSERT INTO 課程 (課程名稱,課程代號,授課教師,修習年級,必選修,學年,學期,開課學制)  VALUES(?,?,?,?,?,?,?,?)";
				ps = connDB.prepareStatement(sql);
				for(int k=0;k<excelcolumn;k++)
				{
					for(int i=0;i<=excelrow;i++)
					{ 
						if(j<exceldata.size())
						{
							obj[i] = exceldata.get(j);
							ps.setString(k+1, obj[i].toString());
							j++;
						}
					}
					ps.executeUpdate();
				}*/
				
				// 0 學年 1 學期 2 選課代碼 3 開課學制 4 授課教師 5 修習年級 6 必選修 7 學分數 8 課程名稱
				i=0;
				while(i<total_class.size())
				{					
					String sql_judge = "INSERT INTO 課程 (課程名稱,課程代碼,授課教師,修習年級,必選修,學分數,學年,學期,開課學制)  "
										+ "select '"+total_class.get(i+8)+"', "
										+ "'"+total_class.get(i+2)+"' , "
										+ "'"+total_class.get(i+4)+"',"
										+ "'"+total_class.get(i+5)+"' , "
										+ "'"+total_class.get(i+6)+"' , "
										+ "'"+total_class.get(i+7)+"' , "
										+ "'"+total_class.get(i)+"' , "
										+ "'"+total_class.get(i+1)+"' , "
										+ "'"+total_class.get(i+3)+"' "
										+ "from dual where not exists ( select * from 課程 "
										+ " where 課程代碼 = '"+total_class.get(i+2)+"' "
										+ "and 學年 = '"+total_class.get(i)+"' "
										+ "and 學期 = '"+total_class.get(i+1)+"')";
					ps = connDB.prepareStatement(sql_judge);
					ps.executeUpdate();
					i+=9;
				}
				//學年 = 1  學期 = 2  課程代碼 = 3   開課學制 = 4  授課教師 = 5   修習年級 = 6  必選修 = 7  學分數 = 8  課程名稱 = 12
				for(k=0;k<excelrow;k++)
				{
					String sql_judge = "INSERT INTO 學生資料 (學年,學期,必選修,學號,姓名,成績,課程名稱,課程代碼,修課年級) "
									+ "select '"+exceldata.get(0)+"' , "
									+ "'"+exceldata.get(1)+"' , "
									+ "'"+exceldata.get(6)+"' , "
									+ "'"+exceldata.get(8)+"' , "
									+ "'"+exceldata.get(9)+"' , "
									+ "'"+exceldata.get(10)+"' , "
									+ "'"+exceldata.get(11)+"' , "
									+ "'"+exceldata.get(2)+"' , "
									+ "'"+exceldata.get(12)+"'"
									+ "from dual where not exists ( select * from 學生資料 "
									+ "where 學號 = '"+exceldata.get(8)+"' "
									+ "and 課程代碼 = '"+exceldata.get(2)+"' "
									+ "and 修課年級 = '"+exceldata.get(12)+"')";
					ps = connDB.prepareStatement(sql_judge);
					ps.executeUpdate();
					for(int num = 0; num<13; num++)
					{
						exceldata.remove(0);
					}
				}
				ps.close();
				st.close();
				connDB.close();
		}catch(ClassNotFoundException e){
			System.out.print("error1");
		}catch(SQLException e){
			System.out.print("error2");
		}
		
	}
	
	
	public static void main(String args[]) throws IOException, SQLException
	{
		gui();
	}
}


