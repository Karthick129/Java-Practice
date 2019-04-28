package org.univ;

public class Stringarray1 {

	public static void main(String[] args) {
		String text = "12344!@#$THILLAIarasu";
		int x = text.length();
		String lcc= "";
		String upp= "";
		String dig= "";
		String sc= "";
		for (int i = 1; i < x; i++)
		{
			char ch = text.charAt(i);
			if (Character.isLowerCase(ch))
			{
				lcc = lcc + ch;
				
			}
		}
			char lc[] = new char[lcc.length()];
			for (int j = 0; j < lcc.length(); j++)
			{
				lc[j] = lcc.charAt(j);
			}
			for (char c : lc) {
				System.out.println("lower case is    " + c);

		}
			
			for (int i = 1; i < x; i++)
			{
				char ch = text.charAt(i);
				if (Character.isUpperCase(ch))
				{
					upp = upp + ch;
					
				}
			}
				char up[] = new char[upp.length()];
				for (int j = 0; j < upp.length(); j++)
				{
					up[j] = upp.charAt(j);
				}
				for (char c : up) {
					System.out.println("upper case is   " + c);

			}
				
				for (int i = 1; i < x; i++)
				{
					char ch = text.charAt(i);
					if (Character.isDigit(ch))
					{
						dig = dig + ch;
						
					}
				}
					char di[] = new char[dig.length()];
					for (int j = 0; j < dig.length(); j++)
					{
						di[j] = dig.charAt(j);
					}
					for (char c : di) {
						System.out.println("Digit is   " + c);

				}
					
					
					for (int i = 1; i < x; i++)
					{
						char ch = text.charAt(i);
						if (Character.isLowerCase(ch))
						{
						}
						else if(Character.isUpperCase(ch))
						{
						}
						else if(Character.isDigit(ch))
						{
						}
						else 
						{
							sc=sc+ch;
						}
					}
					char scc[] = new char[sc.length()];
					for (int j = 0; j < sc.length(); j++)
					{
						scc[j] = sc.charAt(j);
					}
					for (char c : scc) {
						System.out.println("Special character is   " + c);

				}
					
					

					
					
				}



}
