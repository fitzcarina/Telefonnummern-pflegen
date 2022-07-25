using System;

public class Helper_DB
{
	//public Helper_DB()
	//{
	//}

    public static string db_connection()
    {
        string db_string = @"server=vmsql01\prod;database=schnupp; trusted_connection=yes";
        return db_string;
    }
}
