package org.univ;

public class Stringarray2 {

	public static void main(String[] args) {
		String text = "12344!@#$THILLAIarasu";
		int x = text.length();
		char lc[] = new char[x];
		char uc[] = new char[x];
		char di[] = new char[x];
		char sc[] = new char[x];
		for (int i = 1; i < x; i++)
		{
			char ch = text.charAt(i);
			if (Character.isLowerCase(ch))
			{

				lc[i] = text.charAt(i);
				
			}
			else if(Character.isUpperCase(ch))
			{
				uc[i] = text.charAt(i);
			}
			else if(Character.isDigit(ch))
			{
				di[i] = text.charAt(i);
			}
			else 
			{
				sc[i] = text.charAt(i);
			}
		}
			
			for (char c : lc) {
				System.out.println("lower case is    " + c);
		}
			for (char c : uc) {
				System.out.println("upper case is    " + c);
		}
			for (char c : di) {
				System.out.println("digit is    " + c);
		}
			for (char c : sc) {
				System.out.println("special character is    " + c);
		}
			
	
	}

}
