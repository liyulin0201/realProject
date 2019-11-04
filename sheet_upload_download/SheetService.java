package com.ccblife.swp.modules.common.service;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.http.HttpResponse;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.multipart.MultipartFile;


//表单操作
public interface SheetService {
	//导入codeMapping excel数据
	Integer importCodeMappingExcel(MultipartFile myFile) throws Exception;
	//导入sourceDataCode excel数据
	Integer importSourceDataCodeExcel(MultipartFile myFile) throws Exception;
	//导入AimDataCode excel数据
	Integer importAimDataCodeExcel(MultipartFile myFile) throws Exception;
	//导出codeMapping excel数据
	void exportCodeMappingExcel(List<String> ids,HttpServletResponse response) throws IOException;
	//导出SourceData excel数据
	void exportSourceDataExcel(List<String> ids,HttpServletResponse response) throws IOException;
	////导出AimData excel数据
	void exportAimDataExcel(List<String> ids,HttpServletResponse response) throws Exception;
	//模板下载
	void downLoadExcel(HttpServletRequest req, HttpServletResponse res);
}
