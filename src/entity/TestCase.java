package entity;

public class TestCase {
	String excelNo;//excel行号
	String tcName;//用例名称
	String tcNum;//用例编号
	String tcAuthor;//用例作者
	String tcLevel;//用例级别
	String tcVersion;//用例版本
	String tcPrecondition;//用例前置条件
	String tcVersionRecord;//版本更新记录
	String tcApplyProduct;//使用产品
	String tcAutoTestTime;//自动测试时间
	String tcStep;//用例步骤
	String tcResult;//预期结果
	
	
	public String getExcelNo() {
		return excelNo;
	}
	public TestCase setExcelNo(String excelNo) {
		this.excelNo = excelNo;
		return this;
	}
	public String getTcName() {
		return tcName;
	}
	public TestCase setTcName(String tcName) {
		this.tcName = tcName;
		return this;
	}
	public String getTcNum() {
		return tcNum;
	}
	public TestCase setTcNum(String tcNum) {
		this.tcNum = tcNum;
		return this;
	}
	public String getTcAuthor() {
		return tcAuthor;
	}
	public TestCase setTcAuthor(String tcAuthor) {
		this.tcAuthor = tcAuthor;
		return this;
	}
	public String getTcLevel() {
		return tcLevel;
	}
	public TestCase setTcLevel(String tcLevel) {
		this.tcLevel = tcLevel;
		return this;
	}
	public String getTcVersion() {
		return tcVersion;
	}
	public TestCase setTcVersion(String tcVersion) {
		this.tcVersion = tcVersion;
		return this;
	}
	public String getTcPrecondition() {
		return tcPrecondition;
	}
	public TestCase setTcPrecondition(String tcPrecondition) {
		this.tcPrecondition = tcPrecondition;
		return this;
	}
	public String getTcVersionRecord() {
		return tcVersionRecord;
	}
	public TestCase setTcVersionRecord(String tcVersionRecord) {
		this.tcVersionRecord = tcVersionRecord;
		return this;
	}
	public String getTcApplyProduct() {
		return tcApplyProduct;
	}
	public TestCase setTcApplyProduct(String tcApplyProduct) {
		this.tcApplyProduct = tcApplyProduct;
		return this;
	}
	public String getTcAutoTestTime() {
		return tcAutoTestTime;
	}
	public TestCase setTcAutoTestTime(String tcAutoTestTime) {
		this.tcAutoTestTime = tcAutoTestTime;
		return this;
	}
	public String getTcStep() {
		return tcStep;
	}
	public TestCase setTcStep(String tcStep) {
		this.tcStep = tcStep;
		return this;
	}
	public String getTcResult() {
		return tcResult;
	}
	public TestCase setTcResult(String tcResult) {
		this.tcResult = tcResult;
		return this;
	}
	
	public static TestCase build() {
		return new TestCase();
	}
	@Override
	public String toString() {
		return "TestCase [excelNo=" + excelNo + ", tcName=" + tcName + ", tcNum=" + tcNum + ", tcAuthor=" + tcAuthor
				+ ",tcLevel=" + tcLevel + ", tcVersion=" + tcVersion + ", tcPrecondition=" + tcPrecondition + ", tcVersionRecord="
				+ tcVersionRecord + ", tcApplyProduct=" + tcApplyProduct + ", tcAutoTestTime=" + tcAutoTestTime
				+ ", tcStep=" + tcStep + ", tcResult=" + tcResult + "]";
	}
	
	
}
