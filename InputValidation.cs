namespace Cut__copy_and_paste
{
    public static class InputValidation
    {
        public static string UserValidationInput(string UserInput)
        {
            //Cuts the empty space in front and at the end of the string then capitalizes the first character when the user types in a word
            if (!string.IsNullOrEmpty(UserInput) && !string.IsNullOrWhiteSpace(UserInput))
            {
                UserInput = UserInput.Trim();
                return UserInput.Substring(0, 1).ToUpper() + UserInput.Substring(1);
            }

            //If no input is entered or left blank, this field is saved as null
            //if (string.IsNullOrEmpty(UserInput) || string.IsNullOrWhiteSpace(UserInput))
            else
            {
                return null;
            }
        }
    }
}