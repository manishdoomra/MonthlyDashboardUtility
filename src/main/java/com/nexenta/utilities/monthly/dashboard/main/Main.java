package com.nexenta.utilities.monthly.dashboard.main;

import java.util.Properties;

import org.springframework.context.ApplicationContext;
import org.springframework.context.support.ClassPathXmlApplicationContext;

import com.nexenta.utilities.monthly.dashboard.generator.IMonthlyDashboardUtility;
import com.nexenta.utilities.monthly.dashboard.generator.MonthlyDashboardUtilityImpl;

public class Main {
	
	ApplicationContext appContext;
	
	
	public Main(){
		appContext = new ClassPathXmlApplicationContext("spring.xml");
		init();
	}
	
	private void init(){
		Properties properties = appContext.getBean("myproperties", Properties.class);
		IMonthlyDashboardUtility utility  = appContext.getBean("monthlyDashboardUtility", MonthlyDashboardUtilityImpl.class);
		utility.generateMonthlyDashboardTestRailData(properties.getProperty(TESTRAIL_RAW_FILE_PROPERTY), properties.getProperty(TESTRAIL_SECTIONS_AND_ITS_CATEGORIES_FILE_PROPERTY), properties.getProperty(GENERATED_DASHBOARD_FILE_PROPERTY));		
	}
	
	private static final String TESTRAIL_RAW_FILE_PROPERTY = "Testrail.raw.data.excel.file";
	private static final String TESTRAIL_SECTIONS_AND_ITS_CATEGORIES_FILE_PROPERTY = "Section.Categories.file";
	private static final String GENERATED_DASHBOARD_FILE_PROPERTY = "Dashboard.Generated.Excel.file";
			
	public static void main(String[] args) {
		new Main();	

	}

}
