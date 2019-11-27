package gui;

import java.awt.Color;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import javax.swing.BorderFactory;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import entity.TestCase;
import util.ExcelUtil;

public class CFrame extends JFrame implements ActionListener{
	
	private static final long serialVersionUID = 1L;
	
	//定义界面元素
	//定义label
	JLabel im_label = new JLabel("用例文件路径:");    	
	JLabel ex_label = new JLabel("脚本框架存储格式:");
	JLabel result_label = new JLabel("脚本框架生成日志:");
	
	//输入框
	JTextField txt_impath = new JTextField(300);
	JScrollPane scroll = new JScrollPane();
	JTextArea textArea_expath = new JTextArea(""); 	
	//定义button
	ButtonGroup ex_group = new ButtonGroup();
	JButton im_btn = new JButton("选择用例文件");
    JButton ex_btn = new JButton("生成脚本框架");
    //文件选择弹窗
    JFileChooser chooser = new JFileChooser(new File("."));
    //单选按钮
    JRadioButton[] ex_radio = new JRadioButton[5];
    //读入文件
    File imfile = null;
    //输出文件
    File exfile = null;
    //单选框选择值
    String ex_format = null;
    
    //加载定义的元素并设定对应样式
    CFrame(){        
    	im_label.setBounds(10, 10, 180, 30);
        add(im_label);         
        ex_label.setBounds(10, 80, 180, 30);
        add(ex_label);
        result_label.setBounds(10, 200, 180, 30);
        add(result_label);
        txt_impath.setBounds(10, 40, 320, 30);
        txt_impath.setEditable(false);
        add(txt_impath); 
        //textArea日志输出框
        textArea_expath.setBounds(10, 230, 320, 100);
        textArea_expath.setLineWrap(true);
        textArea_expath.setBorder(BorderFactory.createLineBorder(Color.blue));      
        scroll.setBounds(10, 230, 320, 100);
        scroll.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
        scroll.setViewportView(textArea_expath);
        add(scroll); 
        //添加按钮
        im_btn.addActionListener(this);
        im_btn.setBounds(10, 140, 150, 50);
        add(im_btn);  
        ex_btn.addActionListener(this);
        ex_btn.setBounds(180, 140, 150, 50);
        add(ex_btn);      
        //添加单选框
        ex_radio[0] = new JRadioButton("*.java",true); 
        ex_radio[1] = new JRadioButton("*.sh");
        ex_radio[2] = new JRadioButton("*.txt");
        ex_radio[3] = new JRadioButton("*.py");
        ex_radio[4] = new JRadioButton("*.c"); 
        for(int i = 0; i < ex_radio.length; i++) {
        	ex_radio[i].setBounds(10+60*i, 105, 60, 30); 
        	ex_radio[i].setBackground(Color.white);
        	add(ex_radio[i]);
        	ex_group.add(ex_radio[i]); 
        }
        //设置整体样式
        setTitle("脚本框架生成工具 V1.4.1");
        setBounds(360,360,450,400);  
    	setLayout(null);
    	setVisible(true);
    	getContentPane().setBackground(Color.white);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);        
    }
    
    public void actionPerformed(ActionEvent e) {
    	//选择用例文件按钮事件
    	if(e.getSource() == im_btn) {
    		 chooser.setDialogTitle("打开测试用例文档");
    		 chooser.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
    	    int press = chooser.showOpenDialog(null);
    		if(press == JFileChooser.APPROVE_OPTION) {
    			imfile = chooser.getSelectedFile();
    			txt_impath.setText(imfile.getPath()); 
    		}	
    		else
    			txt_impath.setText("");     		
    	}
    	//生成脚本框架按钮事件
    	if(e.getSource()==ex_btn) {     		
    			try {
    				  chooser.setDialogTitle("设定脚本保存路径");
    				  chooser.setFileSelectionMode(JFileChooser.SAVE_DIALOG|JFileChooser.DIRECTORIES_ONLY );
    				  int press = chooser.showSaveDialog(null);
    	    		  if(press==JFileChooser.APPROVE_OPTION) {
    	    			  exfile=chooser.getSelectedFile();
    	    			  textArea_expath.setText("");
    	    			  if(ex_radio[1].isSelected()==true) {
    	    				  ex_format=".sh";
    	    			  }else if(ex_radio[2].isSelected()==true){
    	    				  ex_format=".txt";
    	    			  }else if(ex_radio[3].isSelected()==true){
    	    				  ex_format=".py";
    	    			  }else if(ex_radio[4].isSelected()==true){
    	    				  ex_format=".c";
    	    			  }else{
    	    				  ex_format=".java";
    	    			  }
    	    			  //读取Excel数据到List
    	    			  if (imfile == null) {
    	    				  textArea_expath.append("请选择用例文件!\n");
    	    			  } else {
    	    				  List<TestCase> caseList = ExcelUtil.loadExcel(imfile);
        	    			  textArea_expath.append("共计读取: " + caseList.size() + "条用例信息.\n");
        	    			  //遍历List生成脚本
        	    			  
        	    			  Map<String, String> resultMap = ExcelUtil.toScript(caseList, exfile.getPath(), ex_format);
        	    			  String msg = resultMap.get("errorMsg");
        	    			  if (!ExcelUtil.isEmpty(msg)) {
        	    				  textArea_expath.append("如下用例存在问题请确认:\n" + msg);
        	    			  } else {
        	    				  textArea_expath.append("脚本框架已经输出完成！！\n");
        	    			  }
        	    			  textArea_expath.append("共计输出: " + resultMap.get("successNo") + "个脚本.\n");
    	    			  }
    	    		  }
				} catch (IllegalArgumentException | InvalidFormatException | IOException e1) {	
					e1.printStackTrace();
					textArea_expath.append(e1.getMessage());
				}  catch (Exception e2) {	
					textArea_expath.append(e2.getMessage());
					e2.printStackTrace();
				}			
    	}    	
    }
    
    @SuppressWarnings("unused")
	public static void main(String[] args){    	
    	CFrame frame=new CFrame();  
    } 
}
