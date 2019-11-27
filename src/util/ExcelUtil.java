package util;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.swing.JTextArea;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import entity.TestCase;
/**
 * @version 1.0
 * @name Excel常用方法工具类
 * @describe Excel常用方法合集
 * @author zhangsong01
 *
 */
@SuppressWarnings("unused")
public class ExcelUtil {
	
	private static final int tc_none=0; //层级
	private static final int tc_name=4; //测试用例名称
	private static final int tc_num=6; //测试用例编号
	private static final int tc_precond=7; //测试准备前提条件
	private static final int tc_tstep=13; //测试步骤
	private static final int tc_tresult=14; //预期结果
	private static final int tc_ver=17; //测试用例版本
	private static final int tc_chver=18; //版本变更记录
	private static final int tc_author=19; //用例作者
	private static final int tc_level=23; //用例级别
	private static final int tc_product=25; //适用产品
	private static final int tc_autotime=29; //自动测试时间
	/**
	 * 读取Excel中的用例信息
	 * @param file
	 * @return List<TestCase>
	 * @throws Exception
	 */
	public static List<TestCase> loadExcel(File file) throws Exception, IOException {
		List<TestCase> testCaseList = new ArrayList<>();
		int excelNo = 0;
		try {
			if (!file.exists()) {   
    			throw new IOException("未找到文件,请确认文件路径!");            
    		} 
			String fileName = file.getName();
			String fileType = fileName.substring(fileName.lastIndexOf("."), fileName.length());
    		if (!".xlsx".equals(fileType) && !".xls".equals(fileType)) {
    			System.out.println(fileType);
    			throw new IOException("文件类型错误,请选择Excel文件！");
    		}
			FileInputStream imstream = new FileInputStream(file);  
			Workbook workbook = WorkbookFactory.create(imstream); 
    		 
			Sheet sheet = workbook.getSheetAt(0);
			int lastRowIndex = sheet.getLastRowNum();
			//验证表头
			if (!checkExcelHeader(sheet)) {
				throw new IOException("文件模板有误，请确认文件模板！");
			}
    		for (int idx = 3; idx <= lastRowIndex; idx++) {
    			excelNo = idx + 1;
    			Row row = sheet.getRow(idx);
				// 如果该行为层级结构行，跳过下面判断
    			
				if (row.getCell(tc_name) == null) {
					continue;
				}
				String tcName = row.getCell(tc_name).toString();//空指针
				if (isEmpty(tcName)) {
					continue;
				}
				String tcNum = null;
				if (null != row.getCell(tc_num) && !("").equals(row.getCell(tc_num).toString())) {
					tcNum = row.getCell(tc_num).toString();
				}
				String tcPrecond = null;
				if (null != row.getCell(tc_precond) && !("").equals(row.getCell(tc_precond).toString())) {
					tcPrecond = row.getCell(tc_precond).toString();
				}
				String tcTstep = null;
				if (null != row.getCell(tc_tstep) && !("").equals(row.getCell(tc_tstep).toString())) {
					tcTstep = row.getCell(tc_tstep).toString();
				}
				String tcTresult = null;
				if (null != row.getCell(tc_tresult) && !("").equals(row.getCell(tc_tresult).toString())) {
					tcTresult = row.getCell(tc_tresult).toString();
				}
				String tcVer = null;
				if (null != row.getCell(tc_ver) && !("").equals(row.getCell(tc_ver).toString())) {
					tcVer = row.getCell(tc_ver).toString();
				}
				String tcChver = null;
				if (null != row.getCell(tc_chver) && !("").equals(row.getCell(tc_chver).toString())) {
					tcChver = row.getCell(tc_chver).toString();
				}
				String tcAuthor = null;
				if (null != row.getCell(tc_author) && !("").equals(row.getCell(tc_author).toString())) {
					tcAuthor = row.getCell(tc_author).toString();
				}
				String tcLevel = null;
				if (null != row.getCell(tc_level) && !("").equals(row.getCell(tc_level).toString())) {
					tcLevel = row.getCell(tc_level).toString();
				}
				String tcProduct = null;
				if (null != row.getCell(tc_product) && !("").equals(row.getCell(tc_product).toString())) {
					tcProduct = row.getCell(tc_product).toString();
				}
				String tcAutotime = null;
				if (null != row.getCell(tc_autotime) && !("").equals(row.getCell(tc_autotime).toString())) {
					tcAutotime = row.getCell(tc_autotime).toString();
				}
				
				TestCase tc = TestCase.build()
								.setExcelNo(String.valueOf(excelNo))
								.setTcName(removeExcelEnter(tcName))
								.setTcNum(tcNum)
								.setTcPrecondition(tcPrecond)
								.setTcStep(tcTstep)
								.setTcResult(tcTresult)
								.setTcVersion(tcVer)
								.setTcVersionRecord(tcChver)
								.setTcAuthor(tcAuthor)
								.setTcLevel(tcLevel)
								.setTcApplyProduct(tcProduct)
								.setTcAutoTestTime(tcAutotime);
				testCaseList.add(tc);
    		}
		} catch (RuntimeException e) {
			System.out.println(e.getMessage());
			throw e;
		} catch (IOException e) {
			System.out.println(e.getMessage());
			throw e;
		} catch (Exception e) {
			System.out.println(e.getMessage());
			throw new Exception("Excel中第"+excelNo+"行,存在错误, 请确认!");
		}
		return testCaseList;
	}
	/**
	 * 校验导入文件表头格式是否正确
	 * @param sheet
	 * @return
	 */
	public static boolean checkExcelHeader(Sheet sheet) {
		try {
			Row headRow = sheet.getRow(1);
			String tcName = headRow.getCell(tc_name).toString();
			String tcNum = headRow.getCell(tc_num).toString();
			String tcPrecond = headRow.getCell(tc_precond).toString();
			String tcTstep = headRow.getCell(tc_tstep).toString();
			String tcTresult = headRow.getCell(tc_tresult).toString();
			String tcVer = headRow.getCell(tc_ver).toString();
			String tcChver = headRow.getCell(tc_chver).toString();
			String tcAuthor = headRow.getCell(tc_author).toString();
			String tcLevel = headRow.getCell(tc_level).toString();
			String tcProduct = headRow.getCell(tc_product).toString();
			String tcAutotime = headRow.getCell(tc_autotime).toString();
			if (!tcName.equals("测试用例名称") || !tcNum.equals("测试用例编号") || !tcPrecond.equals("测试准备-前提条件")
					|| !tcTstep.equals("测试步骤") || !tcTresult.equals("预期结果") || !tcVer.equals("测试用例版本") 
					|| !tcChver.equals("版本变更记录") || !tcAuthor.equals("用例作者") || !tcLevel.equals("测试级别-测试类型")
					|| !tcProduct.equals("适用产品") || !tcAutotime.equals("自动测试时间")) {
				return false;
			}
		} catch (Exception e) {
			return false;
		}
		return true;
	}
	
	/**
	 * 校验空字符串
	 * @param str
	 * @return
	 */
	public static boolean isEmpty(String str){  
        if(str == null || "".equals(str.trim())){  
            return true;  
        }else{  
            return false;  
        }  
    }  
	
	/**
	 * 去掉EXCEl中读取出来的换行
	 * @param str
	 * @return
	 */
	public static String removeExcelEnter(String str) {
		for (int i = 10; i < 14; i++) {
			str = str.replaceAll(String.valueOf((char)i), ";");
		}
		return str;
	}
	/**
	 * 去掉EXCEl中读取出来的换行替换为";"
	 * @param str
	 * @return
	 */
	public static String replaceExcelEnter(String str) {
		for (int i = 10; i < 14; i++) {
			str = str.replaceAll(String.valueOf((char)i), ";");
		}
		return str;
	}
	
	/**
	 * 按照{1.|2.|3.}的格式来拆分字符串
	 * @param str
	 * @return
	 */
	public static String[] splitStep(String str) {
		String regex = "[0-9]{1,2}\\.";
		str = replaceExcelEnter(str);
		String[] strs = str.split(regex);
		return strs;
	}

	/**
	 * 
	 * @param src_file
	 * @param dest_filepath
	 * @param file_format
	 * @param jTextArea
	 * @throws Exception
	 */
	public static Map<String, String> toScript(List<TestCase> caseList, String dest_filepath, String file_format) throws IOException, Exception {			
		String ex_comment = "";
		String testCaseName = "";
		String errorMsg = "";
		String tab = "    ";//自定义4个空格
		Map<String, String> resultMap = new HashMap<>();
		List<String> numList = new ArrayList<>();
		int successNo = 0;
        try { 	
    		if (file_format==".sh"||file_format==".py") {
    			ex_comment="#";
    		} else {
    			ex_comment="//";
    		}
			for(TestCase tc : caseList) { 
				if ( tc.getTcAuthor() == null  ||tc.getTcAutoTestTime() == null || tc.getTcName() == null 
						|| tc.getTcNum() == null || tc.getTcPrecondition() == null 
						|| tc.getTcResult() == null || tc.getTcStep() == null ) {
					errorMsg += "第" + tc.getExcelNo()+"行,用例名: " + tc.getTcName() + "\n ";
					continue;
				}
				successNo++;
				testCaseName = tc.getTcName();
				numList.add(tc.getTcNum());
				String dest_filename=dest_filepath+'\\'+tc.getTcNum()+file_format;
				BufferedWriter fOut = new BufferedWriter(new FileWriter(dest_filename));
				fOut.append("package SCRIPT;");
				fOut.newLine();
				fOut.newLine();
				fOut.append(ex_comment + "-----------文件导入部分---------------");
				fOut.newLine();
				fOut.append("import BD.Public.*;");
				fOut.newLine();
				fOut.newLine();
				//用例信息(用例表中获取)
				fOut.append(ex_comment + "--------------------------用例信息-------------------------------");
				fOut.newLine();
				fOut.append("/**");
				fOut.newLine();
				fOut.append(" * 用例名称： " + tc.getTcName());
				fOut.newLine();
				fOut.append(" * 用例编号：" + tc.getTcNum());
				fOut.newLine();
				fOut.append(" * 用例版本：" + tc.getTcVersion());
				fOut.newLine();
				fOut.append(" * 版本变更记录：" + tc.getTcVersionRecord());
				fOut.newLine();
				fOut.append(" * 用例作者：" + tc.getTcAuthor());
				fOut.newLine();
				fOut.append(" * 测试级别-测试类型：" + tc.getTcLevel());
				fOut.newLine();
				fOut.append(" * 预置条件：" + ExcelUtil.removeExcelEnter(tc.getTcPrecondition()));
				fOut.newLine();
				fOut.append(" * 适用产品：" + tc.getTcApplyProduct());
				fOut.newLine();
				fOut.append(" * 自动化测试时间：" + tc.getTcAutoTestTime());
				fOut.newLine();
				fOut.append(" */");
				fOut.newLine();
				fOut.newLine();
				//脚本信息(只生成选项)
				fOut.append(ex_comment + "--------------------------脚本信息-------------------------------");
				fOut.newLine();
				fOut.append("/**");
				fOut.newLine();
				fOut.append(" * @author ");
				fOut.newLine();
				fOut.append(" * @time ");
				fOut.newLine();
				fOut.append(" * @version ");
				fOut.newLine();
				fOut.append(" * @environment ");
				fOut.newLine();
				fOut.append(" * @description ");
				fOut.newLine();
				fOut.append(" * @change ");
				fOut.newLine();
				fOut.append(" */");
				fOut.newLine();
				fOut.newLine();
				//主类开始(脚本名字为用例编号)
				fOut.append("public class " + tc.getTcNum() + " {");
				fOut.newLine();
				//变量部分(用例有几个步骤就包含几个FAIL)
				String[] tcSteps = ExcelUtil.splitStep(tc.getTcStep());
				String allStepResult = tab + "private static String[] allStepResult = {";
				for (int i = 1; i < tcSteps.length; i++) {
					allStepResult += "\"FAIL\", ";
				}
				if (tcSteps.length > 1) {
					allStepResult = allStepResult.substring(0, allStepResult.length()-2);
				}
				allStepResult += " };";
				
				fOut.append(tab + ex_comment + "--------------------------变量部分-------------------------------");
				fOut.newLine();
				fOut.append(tab + "private static Machine machine = BD.Public.Config.getMachineObjectByType(\"SUT\");");
				fOut.newLine();
				fOut.append(tab + "private static String OSIP = machine.getOSIP();");
				fOut.newLine();
				fOut.append(tab + "private static String OSUserName = machine.getOSUserName();");
				fOut.newLine();
				fOut.append(tab + "private static String OSPassword = machine.getOSPassword();");				
				fOut.newLine();
				fOut.newLine();
				fOut.append(tab + "private static TestData testData = BD.Public.Config.getTestDataObjectByName();");
				fOut.newLine();
				fOut.newLine();
				fOut.append(tab + "private static String blankSpaces=\"                       \";");
				fOut.newLine();
				fOut.newLine();
				fOut.append(allStepResult);
				
				// 脚本主体
				fOut.newLine();
				fOut.newLine();
				fOut.append(tab + ex_comment + "--------------------------脚本主体-------------------------------");
				fOut.newLine();
				fOut.append(tab + "public static void main(String[] args) throws Exception {");
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces + \"该脚本所使用的配置文件中的参数如下：\");");
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces + \"OSIP：\" + OSIP);");
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces +\"OSUserName：\" + OSUserName);");
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces + \"OSPassword：\" + OSPassword);");
				fOut.newLine();
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces + \"该脚本所使用的测试数据如下：\");");
				fOut.newLine();
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(\"===========用例名称: "+ tc.getTcName() + "===========\");");
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(\"===========用例编号: "+ tc.getTcNum() + "===========\");");
				fOut.newLine();
				fOut.newLine();
				
				// 测试机器环境检查
				fOut.append(tab + tab + ex_comment + "测试环境检查 ");
				fOut.newLine();
				fOut.append(tab + tab + "boolean machineStatus = BD.Public.Environment.testMachineEnvironment(OSIP, OSUserName, OSPassword);");
				fOut.newLine();
				fOut.append(tab + tab + "if (!machineStatus) {");
				fOut.newLine();
				fOut.append(tab + tab + tab + "System.out.println(blankSpaces +  \"测试环境异常，脚本无法执行，测试FAIL\");");
				fOut.newLine();
				fOut.append(tab + tab + tab + "System.exit(255);");
				fOut.newLine();
				fOut.append(tab + tab + tab + "return;");
				fOut.newLine();
				fOut.append(tab + tab + "}");
				
				// 测试开始
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces + \"———————————————————测试开始———————————————————\");");
				String[] tcResults = {};
				if (tc.getTcResult() != null && tc.getTcResult() != "") {
					tcResults = ExcelUtil.splitStep(tc.getTcResult());
				}
				for (int i = 1; i < tcSteps.length; i++) {
					fOut.newLine();
					fOut.newLine();
					fOut.append(tab + tab + "System.out.println(blankSpaces + \"[测试步骤" + i + "]." + tcSteps[i] + "\");");
					if (i < tcResults.length) {
						fOut.newLine();
						fOut.append(tab + tab + "System.out.println(blankSpaces + \"[预期结果" + i + "]." + tcResults[i] + "\");");
					}
					fOut.newLine();
					fOut.newLine();
					fOut.append(tab + tab + "/*if (Check.checkResult(  ,  )) {");
					fOut.newLine();
					fOut.append(tab + tab + tab + "allStepResult[" + (i-1) +"] = \"PASS\";");
					fOut.newLine();
					fOut.append(tab + tab + tab + "Log.printLog(\"步骤"+ i +"与预期结果一致！\");");
					fOut.newLine();
					fOut.append(tab + tab + tab + "Log.printLog(\"————————————————步骤" + i + "测试PASS————————————————\\n\");");
					fOut.newLine();
					fOut.append(tab + tab + "} else {");
					fOut.newLine();
					fOut.append(tab + tab + tab + "Log.printLog(\"步骤"+ i +"与预期结果不一致！\");");
					fOut.newLine();
					fOut.append(tab + tab + tab + "Log.printLog(\"————————————————步骤" + i + "测试FAIL————————————————\\n\");");
					fOut.newLine();
					fOut.append(tab + tab + "}*/");
				}
				fOut.newLine();
				// 测试结束
				fOut.newLine();
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces + \"———————————————————测试结束———————————————————\");");
				// 最终结果
				fOut.newLine();
				fOut.newLine();
				fOut.append(tab + tab + "System.out.println(blankSpaces + \"———————————————————最终结果———————————————————\");");
				fOut.newLine();
				fOut.append(tab + tab + "if (Check.checkAllStepResult(allStepResult)) {");
				fOut.newLine();
				fOut.append(tab + tab + tab + "Log.printLog(\"=============用例:" + tc.getTcName() + " 测试PASS=============\");");
				fOut.newLine();
				fOut.append(tab + tab + tab + "System.exit(0);");
				fOut.newLine();
				fOut.append(tab + tab + "} else {");
				fOut.newLine();
				fOut.append(tab + tab + tab + "Log.printLog(\"=============用例:" + tc.getTcName() + " 测试FAIL=============\");");
				fOut.newLine();
				fOut.append(tab + tab + tab + "System.exit(255);");
				fOut.newLine();
				fOut.append(tab + tab + "}");
				//脚本主体结束
				fOut.newLine();
				fOut.append(tab + "}");
				//主类结束
				fOut.newLine();
				fOut.append("}");
    			fOut.flush();
    			fOut.close();	
			}    
			resultMap.put("errorMsg", errorMsg);
			resultMap.put("successNo", String.valueOf(successNo));
			
			//查重方法
			/*Set<String> set = new HashSet<>();
			numList.stream().forEach(p -> {
				set.add(p);
			});
			System.out.println(set.size());*/
        } catch (IOException e) {   
        	e.printStackTrace();             
        	throw e;
        } catch (Exception e) {   
        	e.printStackTrace();             
        	throw new Exception("用例:"+testCaseName+",存在错误, 请确认!");
        } finally {            
        } 
        return resultMap;
	}    

}
