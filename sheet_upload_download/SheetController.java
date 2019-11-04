package com.ccblife.swp.modules.common.controller;

import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.ModelAndView;

import com.ccblife.swp.exception.BusinessException;
import com.ccblife.swp.modules.common.service.SheetService;

@RestController
@RequestMapping("/sheet")
public class SheetController {

	@Autowired
	private SheetService sheetService;

	@RequestMapping(value = "/importSourceDataExcel.do", method = RequestMethod.POST)
	public String importSourceDataExcel(@RequestParam("file") MultipartFile File) {
		ModelAndView modelAndView = new ModelAndView();
		try {
			Integer num = sheetService.importSourceDataCodeExcel(File);
		} catch (Exception e) {
			e.printStackTrace();
			modelAndView.addObject("msg", e.getMessage());
			return "导入失败";
		}
		modelAndView.addObject("msg", "数据导入成功");
		return "导入成功";
	}

	@RequestMapping(value = "/importAimDataExcel.do", method = RequestMethod.POST)
	public String importAimDataExcel(@RequestParam("file") MultipartFile File) {
		ModelAndView modelAndView = new ModelAndView();
		try {
			Integer num = sheetService.importAimDataCodeExcel(File);
		} catch (Exception e) {
			e.printStackTrace();
			modelAndView.addObject("msg", e.getMessage());
			return "导入失败";
		}
		modelAndView.addObject("msg", "数据导入成功");
		return "导入成功";
	}

	@RequestMapping(value = "/importCodeMappingExcel.do", method = RequestMethod.POST)
	public String importExcel(@RequestParam("file") MultipartFile File) {
		ModelAndView modelAndView = new ModelAndView();
		try {
			Integer num = sheetService.importCodeMappingExcel(File);
		} catch (Exception e) {
			e.printStackTrace();
			modelAndView.addObject("msg", e.getMessage());
			return "导入失败";
		}
		modelAndView.addObject("msg", "数据导入成功");
		return "导入成功";
	}

	@GetMapping(value = "/exportSourceDataExcel.do")
	public void exportSourceDataExcel(String ids, HttpServletResponse response) throws BusinessException {
		if(StringUtils.isEmpty(ids)){
			 throw new BusinessException("ids不能为空");
		}
		List<String> list = Arrays.asList(ids.split(","));
		try {
			if (list != null && !list.isEmpty()){
				sheetService.exportSourceDataExcel(list, response);
			}
		} catch (IOException e) {
			throw new BusinessException("导出失败");
		}

	}

	@GetMapping(value = "/exportAimDataExcel.do")
	public void exportAimDataExcel(String ids,HttpServletResponse response) throws BusinessException {
		if(StringUtils.isEmpty(ids)){
			 throw new BusinessException("ids不能为空");
		}
		List<String> list = Arrays.asList(ids.split(","));
		try {
			if (list != null && !list.isEmpty()){
				sheetService.exportAimDataExcel(list, response);
			}
		} catch (Exception e) {
			throw new BusinessException("导出失败");
		}
	}

	@GetMapping(value = "/exportCodeMappingExcel.do")
	public void exportCodeMappingExcel(String ids,HttpServletResponse response) throws BusinessException {
		if(StringUtils.isEmpty(ids)){
			 throw new BusinessException("ids不能为空");
		}
		List<String> list = Arrays.asList(ids.split(","));
		try {
			if (list != null && !list.isEmpty()){
				sheetService.exportCodeMappingExcel(list, response);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	// 模板下载
	@ResponseBody
	@RequestMapping("/downloadExcel.do")
	public void downLoadTemplate(HttpServletRequest req, HttpServletResponse res) {
		sheetService.downLoadExcel(req, res);
	}

}