package Test_Classes;
import io.restassured.path.json.JsonPath;

import java.io.IOException;
import java.time.LocalDateTime;

import org.testng.Assert;

import Common_API_Method.Patch_Api_Method;
import Req_Repository.Patch_Req_Repository;

public class Patch_TC1 {
	
	public static void extractor() throws IOException {
		for(int i=0;i<5;i++) {
			
		int statuscode=Patch_Api_Method.responsestatuscode(Patch_Req_Repository.baseURI(),Patch_Req_Repository.Patch_Resourse(),Patch_Req_Repository.Patch_TC1());
		System.out.println(statuscode);
		
		if(statuscode==200) {
			
			String Responsebody=Patch_Api_Method.Responsebody(Patch_Req_Repository.baseURI(), Patch_Req_Repository.Patch_Resourse(), Patch_Req_Repository.Patch_TC1());
			System.out.println(Responsebody);
			
			String Requestbody =Patch_Req_Repository.Patch_TC1();
			System.out.println(Requestbody);
			
			validator(Requestbody,Responsebody);
			break;}
		else {
			System.out.println("status is invalid");
			
			
		}
		}
	}

	public static void validator(String Requestbody,String Responsebody) {
		
		
		//Parse Req body
		JsonPath jspReq=new JsonPath(Requestbody);
		String Req_name=jspReq.getString("name");
		String Req_job=jspReq.getString("job");
		LocalDateTime currentdate=LocalDateTime.now();
		String expecteddate=currentdate.toString().substring(0,11);
		
		//parse Response body
		JsonPath jspRes=new JsonPath(Responsebody);
		String Res_name=jspRes.getString("name");
		String Res_job=jspRes.getString("job");
		String Res_updatedAt=jspRes.getString("updatedAt");
		Res_updatedAt=Res_updatedAt.substring(0,11);
		
		//validate response body parameter
		Assert.assertEquals(Res_name, Req_name);
		Assert.assertEquals(Res_job, Req_job);
		Assert.assertEquals(Res_updatedAt, expecteddate);
	}
}
