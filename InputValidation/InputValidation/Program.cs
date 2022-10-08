using System;

namespace Cut_copy_and_paste
public static class InputValidation
{
	public static string CallClass()
	{
		InputValidation.UserValidationInput(UserInput);
	}

	public static string UserValidationInput(string UserInput)
	{
		Console.WriteLine("Input word");
		//Cuts the empty space in front and at the end of the string then capitalizes the first character when the user types in a word
		if (!string.IsNullOrEmpty(UserInput) || !string.IsNullOrWhiteSpace(UserInput))
		{
			UserInput = UserInput.Substring(0, 1).ToUpper() + UserInput.Substring(1);
			UserInput = Console.ReadLine();
		}

		//If no input is entered or left blank, this field is saved as null
		else
		{
			UserInput = null;
			UserInput = Console.ReadLine();
		}
	}
}