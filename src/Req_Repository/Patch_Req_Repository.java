package Req_Repository;

import java.io.IOException;
import java.util.ArrayList;

import Common_API_Method.Common_utility_Method;

public class Patch_Req_Repository {
	
	public static String baseURI() {
		String baseURI="https://reqres.in/";
		return baseURI;
		
	}
	
	public static String Patch_Resourse() {
		String Patch_Resourse="api/users/2";
		return Patch_Resourse;
	}
	
	public static String Patch_TC1() throws IOException {
		ArrayList<String> Req_Data=Common_utility_Method.ReadDataExcel("PatchApi", "TC2");
		System.out.println(Req_Data);
		String Req_name=Req_Data.get(0);
		String Req_job=Req_Data.get(1);
		String Requestbody= "{\r\n"
				+ "    \"name\": \"morpheus\",\r\n"
				+ "    \"job\": \"zion resident\"\r\n"
				+ "}";
		return Requestbody;
	}

}
