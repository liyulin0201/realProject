package com.ccblife.swp.modules.common.service.impl;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.apache.ibatis.session.ExecutorType;
import org.apache.ibatis.session.SqlSession;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.mybatis.spring.SqlSessionTemplate;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.web.multipart.MultipartFile;

import com.ccblife.swp.exception.BusinessException;
import com.ccblife.swp.mapper.codemapping.AimDataCodeMapper;
import com.ccblife.swp.mapper.codemapping.CodeMappingCodeMapper;
import com.ccblife.swp.mapper.codemapping.SourceDataCodeMapper;
import com.ccblife.swp.modules.codemapping.dto.CodeMappingQueryContidionDto;
import com.ccblife.swp.modules.codemapping.entity.AimDataCode;
import com.ccblife.swp.modules.codemapping.entity.CodeMapping;
import com.ccblife.swp.modules.codemapping.entity.SourceDataCode;
import com.ccblife.swp.modules.common.service.SheetService;
import com.ccblife.swp.util.UserHolder;

@Service
public class SheetServiceImpl implements SheetService {

	private final static String XLS = "xls";
	private final static String XLSX = "xlsx";
	private static final int INSERT_ROW = 500;

	@Autowired
	private SourceDataCodeMapper sourceDataCodeMapper;
	@Autowired
	private AimDataCodeMapper aimDataCodeMapper;
	@Autowired
	private CodeMappingCodeMapper codeMapper;
	
	public static String getCellValue(Cell cell) {
		String value = "";
		if (cell != null) {
			// 以下是判断数据的类型
			switch (cell.getCellType()) {
			case HSSFCell.CELL_TYPE_NUMERIC: // 数字
				value = cell.getNumericCellValue() + "";
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					Date date = cell.getDateCellValue();
					if (date != null) {
						value = new SimpleDateFormat("yyyy-MM-dd").format(date);
					} else {
						value = "";
					}
				} else {
					value = new DecimalFormat("0").format(cell.getNumericCellValue());
				}
				break;
			case HSSFCell.CELL_TYPE_STRING: // 字符串
				value = cell.getStringCellValue();
				break;
			case HSSFCell.CELL_TYPE_BOOLEAN: // Boolean
				value = cell.getBooleanCellValue() + "";
				break;
			case HSSFCell.CELL_TYPE_FORMULA: // 公式
				value = cell.getCellFormula() + "";
				break;
			case HSSFCell.CELL_TYPE_BLANK: // 空值
				value = "";
				break;
			case HSSFCell.CELL_TYPE_ERROR: // 故障
				value = "非法字符";
				break;
			default:
				value = "未知类型";
				break;
			}
		}
		return value.trim();
	}

	public static boolean isRowEmpty(Row row) {
		for (int c = row.getFirstCellNum(); c < row.getLastCellNum(); c++) {
			Cell cell = row.getCell(c);
			if (cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK)
				return false;
		}
		return true;
	}

	@Override
	public void downLoadExcel(HttpServletRequest req, HttpServletResponse res) {
		try {

			// 获取要下载的模板名称
			String fileName = "testworkbook.xls";
			// 设置要下载的文件的名称
			res.setHeader("Content-disposition", "attachment;fileName=" + fileName);
			// 通知客服文件的MIME类型
			res.setContentType("application/vnd.ms-excel;charset=UTF-8");
			// 获取文件的路径
			String filePath = getClass().getResource("/templates/testworkbook.xls").getPath();
			FileInputStream input = new FileInputStream(filePath);
			OutputStream out = res.getOutputStream();
			byte[] b = new byte[2048];
			int len;
			while ((len = input.read(b)) != -1) {
				out.write(b, 0, len);
			}
			// 修正 Excel在“xxx.xlsx”中发现不可读取的内容。是否恢复此工作薄的内容？如果信任此工作簿的来源，请点击"是"
			res.setHeader("Content-Length", String.valueOf(input.getChannel().size()));
			input.close();
			out.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	@SuppressWarnings("resource")
	@Override
	public Integer importSourceDataCodeExcel(MultipartFile myFile) throws Exception {
		SqlSession session = sqlSessionTemplate.getSqlSessionFactory().openSession(ExecutorType.BATCH, false);
		Integer resultNo = null;
		try {
			Workbook workbook = null;
			String fileName = myFile.getOriginalFilename();
			if (fileName.endsWith(XLS)) {
				// 2003
				workbook = new HSSFWorkbook(myFile.getInputStream());
			} else if (fileName.endsWith(XLSX)) {
				// 2007
				workbook = new XSSFWorkbook(myFile.getInputStream());
			} else {
				throw new Exception("文件不是Excel文件");
			}
			List<SourceDataCode> sourceDataList = new ArrayList<SourceDataCode>();
			Map<String, String> paramMap = new HashMap<String, String>();
			Sheet sheetNum = null;
			// 对Sheet中的每一行进行迭代
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				sheetNum = workbook.getSheetAt(i);
				Row r = null;
				for (int j = 1; j <= sheetNum.getLastRowNum(); j++) {
					r = sheetNum.getRow(j);
					if (!isRowEmpty(r)) {
						SourceDataCode sdc = new SourceDataCode();
						String codeTypeEncode = getCellValue(r.getCell(1));
						String codeTypeName = getCellValue(r.getCell(2));
						String sourceCode = getCellValue(r.getCell(3));
						String sourceCodeName = getCellValue(r.getCell(4));
						String remark = getCellValue(r.getCell(5));
						String dataSource = getCellValue(r.getCell(6));
						String categoryId = getCellValue(r.getCell(7));
						//组装条件查询是否有源数据存在
						paramMap.put("codeTypeEncode", codeTypeEncode);
						paramMap.put("code", sourceCode);
						List<SourceDataCode> resList = sourceDataCodeMapper.queryListSourceDataCode(paramMap);
						if(resList.size()==0){
							sdc.setCodeTypeEncode(codeTypeEncode);
							sdc.setCodeTypeName(codeTypeName);
							sdc.setCode(sourceCode);
							sdc.setSourceDataName(sourceCodeName);
							sdc.setRemark(remark);
							sdc.setDataSource(dataSource);
							sdc.setCategoryId(categoryId);
							Date time = Calendar.getInstance().getTime();
							if (StringUtils.isBlank(sdc.getCreator())) {
								sdc.setCreator(UserHolder.get().getUserAccount());
							}
							sdc.setCreateTime(time);
							if (StringUtils.isBlank(sdc.getModifier())) {
								sdc.setModifier(UserHolder.get().getUserAccount());
							}
							sdc.setModifyTime(time);
							sdc.setEnabled("1");
							sourceDataList.add(sdc);
						}else{
							throw new BusinessException("已存在源数据");
						}
					}
				}
			}
				// 批量插入操作
				if (sourceDataList.size() <= INSERT_ROW) {
					resultNo = session.insert(
							"com.ccblife.swp.mapper.codemapping.SourceDataCodeMapper.insertBatchSourceDataCode",
							sourceDataList);
				} else {
					List<SourceDataCode> list100 = new ArrayList<SourceDataCode>(new Integer(100));
					List<SourceDataCode> listRest = new ArrayList<SourceDataCode>();

					for (int i = 1; i <= (sourceDataList.size() - 1); i++) {
						list100.add(sourceDataList.get(i));
						if (i % 100 == 0) {
							resultNo = session.insert(
									"com.ccblife.swp.mapper.codemapping.SourceDataCodeMapper.insertBatchSourceDataCode",
									list100);
							list100.clear();
						}
						if (i >= sourceDataList.size() - sourceDataList.size() % 100) {
							listRest.add(sourceDataList.get(i));
						}
					}
					resultNo = session.insert(
							"com.ccblife.swp.mapper.codemapping.SourceDataCodeMapper.insertBatchSourceDataCode", listRest);
				}
				session.commit();
		} catch (Exception e) {
			e.printStackTrace();
			session.rollback();
		}finally{
			session.close();
		}
		return resultNo;
	}

	@Override
	public Integer importAimDataCodeExcel(MultipartFile myFile) throws Exception {
		SqlSession session = sqlSessionTemplate.getSqlSessionFactory().openSession(ExecutorType.BATCH, false);
		Integer resultNo = null;
		try {
			Workbook workbook = null;
			String fileName = myFile.getOriginalFilename();
			if (fileName.endsWith(XLS)) {
				// 2003
				workbook = new HSSFWorkbook(myFile.getInputStream());
			} else if (fileName.endsWith(XLSX)) {
				// 2007
				workbook = new XSSFWorkbook(myFile.getInputStream());
			} else {
				throw new Exception("文件不是Excel文件");
			}
			List<AimDataCode> aimDataList = new ArrayList<AimDataCode>();
			Map<String, String> paramMap = new HashMap<String, String>();
			Sheet sheetNum = null;
			// 对Sheet中的每一行进行迭代
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				sheetNum = workbook.getSheetAt(i);
				Row r = null;
				for (int j = 1; j <= sheetNum.getLastRowNum(); j++) {
					r = sheetNum.getRow(j);
					if (!isRowEmpty(r)) {
						AimDataCode adc = new AimDataCode();
						String codeTypeEncode = getCellValue(r.getCell(1));
						String codeTypeName = getCellValue(r.getCell(2));
						String aimCode = getCellValue(r.getCell(3));
						String aimCodeName = getCellValue(r.getCell(4));
						String remark = getCellValue(r.getCell(5));
						String dataSource = getCellValue(r.getCell(6));
						String categoryId = getCellValue(r.getCell(7));
						// 组装条件查询是否目标源数据存在
						paramMap.put("codeTypeEncode", codeTypeEncode);
						paramMap.put("code", aimCode);
						List<AimDataCode> resList = aimDataCodeMapper.queryListAimDataCode(paramMap);
						if (resList.size() == 0) {
							adc.setCodeTypeEncode(codeTypeEncode);
							adc.setCodeTypeName(codeTypeName);
							adc.setCategoryId(categoryId);
							adc.setCode(aimCode);
							adc.setSourceDataName(aimCodeName);
							adc.setDataSource(dataSource);
							adc.setRemark(remark);
							Date time = Calendar.getInstance().getTime();
							if (StringUtils.isBlank(adc.getCreator())) {
								adc.setCreator(UserHolder.get().getUserAccount());
							}
							adc.setCreateTime(time);
							if (StringUtils.isBlank(adc.getModifier())) {
								adc.setModifier(UserHolder.get().getUserAccount());
							}
							adc.setModifyTime(time);
							adc.setEnabled("1");
							aimDataList.add(adc);
						} else {
							throw new BusinessException("已存在目标数据");
						}
					}
				}
			}

				// 批量插入操作
				if (aimDataList.size() <= INSERT_ROW) {
						resultNo = session.insert(
							"com.ccblife.swp.mapper.codemapping.AimDataCodeMapper.insertBatchAimDataCode",
							aimDataList);
				} else {
					List<AimDataCode> list100 = new ArrayList<AimDataCode>(new Integer(100));
					List<AimDataCode> listRest = new ArrayList<AimDataCode>();

					for (int i = 1; i <= (aimDataList.size()-1); i++) {
						list100.add(aimDataList.get(i));
						if (i % 100 == 0) {
							resultNo = session.insert(
									"com.ccblife.swp.mapper.codemapping.AimDataCodeMapper.insertBatchAimDataCode",
									list100);
							list100.clear();
						}
						if (i >= aimDataList.size() - aimDataList.size() % 100) {
							listRest.add(aimDataList.get(i));
						}
					}
					resultNo = session.insert(
							"com.ccblife.swp.mapper.codemapping.AimDataCodeMapper.insertBatchAimDataCode", listRest);
					
				}
				session.commit();
		} catch (Exception e) {
			e.printStackTrace();
			session.rollback();
		} finally {
			session.close();
		}
		return resultNo;
	}
	
	
	@SuppressWarnings({ "resource" })
	@Override
	@Transactional
	public Integer importCodeMappingExcel(MultipartFile myFile) throws Exception {
		SqlSession session = sqlSessionTemplate.getSqlSessionFactory().openSession(ExecutorType.BATCH, false);
		Integer resultNo = null;
		try {
			Workbook workbook = null;
			String fileName = myFile.getOriginalFilename();
			if (fileName.endsWith(XLS)) {
				// 2003
				workbook = new HSSFWorkbook(myFile.getInputStream());
			} else if (fileName.endsWith(XLSX)) {
				// 2007
				workbook = new XSSFWorkbook(myFile.getInputStream());
			} else {
				throw new Exception("文件不是Excel文件");
			}
			List<CodeMapping> codeMappingList = new ArrayList<CodeMapping>();
			Map<String, String> paramMap = new HashMap<String, String>();
			Sheet sheetNum = null;
			// 对Sheet中的每一行进行迭代
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				sheetNum = workbook.getSheetAt(i);
				Row r = null;
				for (int j = 1; j <= sheetNum.getLastRowNum(); j++) {
					r = sheetNum.getRow(j);
					if (!isRowEmpty(r)) {
						CodeMapping cmq = new CodeMapping();
						String codeTypeName = getCellValue(r.getCell(1));
						String codeTypeEncode = getCellValue(r.getCell(2));
						String sourceCode = getCellValue(r.getCell(3));
//						String sourceCodeName = getCellValue(r.getCell(4));
						String aimCode = getCellValue(r.getCell(5));
//						String aimCodeName = getCellValue(r.getCell(6));
						String remark = getCellValue(r.getCell(7));
						cmq.setCodeTypeEncode(codeTypeEncode);// 代码类型编码
						cmq.setCodeTypeName(codeTypeName);// 代码类型名称
						SourceDataCode sourceDataCode = new SourceDataCode();
						AimDataCode aimDataCode = new AimDataCode();
						//拼装查询条件
						sourceDataCode.setCodeTypeEncode(codeTypeEncode);
						sourceDataCode.setCode(sourceCode);
						aimDataCode.setCodeTypeEncode(codeTypeEncode);
						aimDataCode.setCode(aimCode);
						//执行查询
						SourceDataCode ret1 = sourceDataCodeMapper.selectOneSourceDataCode(sourceDataCode);
						AimDataCode ret2 = aimDataCodeMapper.selectOneAimData(aimDataCode);
						if (ret1 != null && ret2 != null) {
							cmq.setSourceId(ret1.getSourceId());// 源数据代码
							cmq.setAimId(ret2.getAimId());// 目标数据代码
						} else {
							throw new BusinessException("请新增源数据或目标数据");
						}
						cmq.setRemark(remark);// 备注
						paramMap.put("sourceId", ret1.getSourceId());
						paramMap.put("aimId", ret2.getAimId());
						// 查询是否有映射存在
						List<CodeMapping> selList = codeMapper.queryListCodeMapping(paramMap);
						paramMap.clear();
						if (selList.size() > 0) {
							throw new BusinessException("已存在映射");
						}else{
							Date time = Calendar.getInstance().getTime();
							if (StringUtils.isBlank(cmq.getCreator())) {
								cmq.setCreator(UserHolder.get().getUserAccount());
							}
							cmq.setCreateTime(time);
							if (StringUtils.isBlank(cmq.getModifier())) {
								cmq.setModifier(UserHolder.get().getUserAccount());
							}
							cmq.setModifyTime(time);
							cmq.setEnabled("1");
							codeMappingList.add(cmq);
						}
					}
				}
			}
				// 批量插入操作
				if (codeMappingList.size() <= INSERT_ROW) {
					resultNo = session.insert(
							"com.ccblife.swp.mapper.codemapping.CodeMappingCodeMapper.insertBatchCodeMapping",
							codeMappingList);
					session.commit();
				} else {
					List<CodeMapping> list100 = new ArrayList<CodeMapping>(new Integer(100));
					List<CodeMapping> listRest = new ArrayList<CodeMapping>();

					for (int i = 1; i <= (codeMappingList.size() - 1); i++) {
						list100.add(codeMappingList.get(i));
						if (i % 100 == 0) {
							resultNo = session.insert(
									"com.ccblife.swp.mapper.codemapping.CodeMappingCodeMapper.insertBatchCodeMapping",
									list100);
							list100.clear();
						}
						if (i >= codeMappingList.size() - codeMappingList.size() % 100) {
							listRest.add(codeMappingList.get(i));
						}
					}
					resultNo = session.insert(
							"com.ccblife.swp.mapper.codemapping.CodeMappingCodeMapper.insertBatchCodeMapping", listRest);
				}
				session.commit();
		} catch (Exception e) {
			e.printStackTrace();
			session.rollback();
		} finally{
			session.close();
		}
		return resultNo;
	}
	
	
	
	@Override
	public void exportSourceDataExcel(List<String> ids,HttpServletResponse response) throws IOException {
		XSSFWorkbook workbook =new XSSFWorkbook(this.getClass().getResourceAsStream("/"+"sourceData.xlsx"));
		XSSFSheet sheet = workbook.getSheetAt(0);
		sheet.setColumnWidth(12,100*25);
		XSSFRow row = sheet.createRow(0);
        
		XSSFCell cell = row.createCell(0);
        cell.setCellValue("目标数据序号");
        cell = row.createCell(1);
        cell.setCellValue("代码类型编码");
        cell = row.createCell(2);
        cell.setCellValue("代码类型名称");
        cell = row.createCell(3);
        cell.setCellValue("目标数据代码");
        cell = row.createCell(4);
        cell.setCellValue("目标数据代码名");
        cell = row.createCell(5);
        cell.setCellValue("备注");
        cell = row.createCell(6);
        cell.setCellValue("数据来源");
        cell = row.createCell(7);
        cell.setCellValue("父级代码");
        
		List<SourceDataCode> list = sourceDataCodeMapper.selectListSourceDataById(ids);
		for (int i = 0; i < list.size(); i++) {
			 row = sheet.createRow(i+1);
			// 第四步，创建单元格，并设置值
			row.createCell(0).setCellValue(list.get(i).getSourceId());
			row.createCell(1).setCellValue(list.get(i).getCodeTypeEncode());
			row.createCell(2).setCellValue(list.get(i).getCodeTypeName());
			row.createCell(3).setCellValue(list.get(i).getCode());
			row.createCell(4).setCellValue(list.get(i).getSourceDataName());
			row.createCell(5).setCellValue(list.get(i).getRemark());
			row.createCell(6).setCellValue(list.get(i).getDataSource());
			if(list.get(i).getCategoryId()!=null){
				row.createCell(7).setCellValue(list.get(i).getCategoryId());
			}
		}
		String filename="源数据代码"+new SimpleDateFormat("MM-dd-HH-mm-ss").format(new Date());
		OutputStream outputStream = response.getOutputStream();
		response.setHeader("Content-disposition", "attachment;filename="
                + new String(filename.getBytes("gb2312") , "ISO8859-1")+".xlsx");//设置文件头编码格式
        response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");//设置类型
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
		
	}

	@Override
	public void exportAimDataExcel(List<String> ids,HttpServletResponse response) throws IOException, BusinessException {
		XSSFWorkbook workbook =new XSSFWorkbook(this.getClass().getResourceAsStream("/"+"aimData.xlsx"));
		XSSFSheet sheet = workbook.getSheetAt(0);
		sheet.setColumnWidth(12,100*25);
		XSSFRow row = sheet.createRow(0);
        
		XSSFCell cell = row.createCell(0);
        cell.setCellValue("目标数据序号");
        cell = row.createCell(1);
        cell.setCellValue("代码类型编码");
        cell = row.createCell(2);
        cell.setCellValue("代码类型名称");
        cell = row.createCell(3);
        cell.setCellValue("目标数据代码");
        cell = row.createCell(4);
        cell.setCellValue("目标数据代码名");
        cell = row.createCell(5);
        cell.setCellValue("备注");
        cell = row.createCell(6);
        cell.setCellValue("数据来源");
        cell = row.createCell(7);
        cell.setCellValue("父级代码");
        
     // 从数据库查询数据
		List<AimDataCode> list = aimDataCodeMapper.selectListAimDataById(ids);
		for (int i = 0; i < list.size(); i++) {
			 row = sheet.createRow(i+1);
			// 第四步，创建单元格，并设置值
			row.createCell(0).setCellValue(list.get(i).getAimId());
			row.createCell(1).setCellValue(list.get(i).getCodeTypeEncode());
			row.createCell(2).setCellValue(list.get(i).getCodeTypeName());
			row.createCell(3).setCellValue(list.get(i).getCode());
			row.createCell(4).setCellValue(list.get(i).getSourceDataName());
			row.createCell(5).setCellValue(list.get(i).getRemark());
			row.createCell(6).setCellValue(list.get(i).getDataSource());
			row.createCell(7).setCellValue(list.get(i).getCategoryId());
			if(list.get(i).getCategoryId()!=null){
				row.createCell(7).setCellValue(list.get(i).getCategoryId());
			}
		}
		String filename="目标数据代码"+new SimpleDateFormat("MM-dd-HH-mm-ss").format(new Date());
		OutputStream outputStream = response.getOutputStream();
		response.setHeader("Content-disposition", "attachment;filename="
                + new String(filename.getBytes("gb2312") , "ISO8859-1")+".xlsx");//设置文件头编码格式
        response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");//设置类型
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
	}

	@Override
	public void exportCodeMappingExcel(List<String> ids,HttpServletResponse response) throws IOException {
		XSSFWorkbook workbook =new XSSFWorkbook(this.getClass().getResourceAsStream("/"+"codeMappingData.xlsx"));
		XSSFSheet sheet = workbook.getSheetAt(0);
		sheet.setColumnWidth(12,100*25);
		XSSFRow row = sheet.createRow(0);
		
		XSSFCell cell = row.createCell(0);
		cell.setCellValue("ID");

		cell = row.createCell(1);
		cell.setCellValue("代码类型名称");

		cell = row.createCell(2);
		cell.setCellValue("代码类型编码");

		cell = row.createCell(3);
		cell.setCellValue("源数据代码");

		cell = row.createCell(4);
		cell.setCellValue("源数据代码名");

		cell = row.createCell(5);
		cell.setCellValue("目标数据代码");

		cell = row.createCell(6);
		cell.setCellValue("目标数据代码名");

		cell = row.createCell(7);
		cell.setCellValue("备注");
		
     // 从数据库查询数据
		List<CodeMappingQueryContidionDto> list = codeMapper.selectListCodeMappingById(ids);
		for (int i = 0; i < list.size(); i++) {
			 row = sheet.createRow(i+1);
			// 第四步，创建单元格，并设置值
			row.createCell(0).setCellValue(list.get(i).getCodeMapping().getPkId());
			row.createCell(1).setCellValue(list.get(i).getCodeMapping().getCodeTypeEncode());
			row.createCell(2).setCellValue(list.get(i).getCodeMapping().getCodeTypeName());
			row.createCell(3).setCellValue(list.get(i).getSourceData().getCode());
			row.createCell(4).setCellValue(list.get(i).getSourceData().getSourceDataName());
			row.createCell(5).setCellValue(list.get(i).getAimDataCode().getCode());
			row.createCell(6).setCellValue(list.get(i).getAimDataCode().getSourceDataName());
			row.createCell(7).setCellValue(list.get(i).getCodeMapping().getRemark());
		}
		String filename="代码映射"+new SimpleDateFormat("MM-dd-HH-mm-ss").format(new Date());
		OutputStream outputStream = response.getOutputStream();
		response.setHeader("Content-disposition", "attachment;filename="
                + new String(filename.getBytes("gb2312") , "ISO8859-1")+".xlsx");//设置文件头编码格式
        response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");//设置类型
        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
	}

	@Autowired
	private SqlSessionTemplate sqlSessionTemplate;
	
}
