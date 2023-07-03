package Test_Classes;

import java.io.IOException;
import java.time.LocalDateTime;

import org.testng.Assert;

import Common_API_Method.Post_Api_Method;
import Req_Repository.Post_Req_Repository;
import io.restassured.path.json.JsonPath;

public class Post_TC1 {
	
	public static void extractor() throws IOException {
		
		for (int i=0;i<5;i++) {
			
			int statuscode=Post_Api_Method.responsestatuscode(Post_Req_Repository.baseURI(),Post_Req_Repository.Post_Resourse(),Post_Req_Repository.Post_TC1());
			System.out.println(statuscode);
			
			if(statuscode==201) {
				
				String Responsebody= Post_Api_Method.Responsebody(Post_Req_Repository.baseURI(),Post_Req_Repository.Post_Resourse(),Post_Req_Repository.Post_TC1());
				System.out.println(Responsebody);
				
				String Requestbody=Post_Req_Repository.Post_TC1();
				
				validator(Requestbody,Responsebody);
				
				break;}
			else {
				System.out.println("Status code is invalid");
			}
		}
	}
	public static void validator(String Requestbody, String Responsebody) {
		
		//parse Request body
		JsonPath jspReq =new JsonPath(Requestbody);
		String Req_name=jspReq.getString("name");
		String Req_job=jspReq.getString("job");
		LocalDateTime currentdate=LocalDateTime.now();
		String expecteddate=currentdate.toString().substring(0,11);
		
		//create object of JsonPath for parse Response body
		JsonPath jspRes=new JsonPath(Responsebody);
		String Res_name=jspRes.getString("name");
		String Res_job=jspRes.getString("job");
		String Res_createdAt=jspRes.getString("createdAt");
		Res_createdAt=Res_createdAt.substring(0,11);
		
		//validate response body parameter
		Assert.assertEquals(Res_name, Req_name);
		Assert.assertEquals(Res_job, Req_job);
		Assert.assertEquals(Res_createdAt, expecteddate);
	}
}
