namespace Schedule.Maker.Models.Helper.TreeStructure
{
    /// <summary>
    ///  Helper class used to establish the right tree structure for the generation.
    /// </summary>
    public abstract class TreeStructureHelper
    {
        /**
         * <summary>
         *  Fuction used to find a given directory using directory's path.
         * </summary>
         * <return>Boolean value that indicates the potential existence of the directory.</return>
         */
        public static bool Find_Directory(string directory_path)
        {
            if (Directory.Exists(directory_path))
                return true;
            return false;
        }

        /**
         * <summary>
         *  Function used to find a given file using file's path.
         * </summary>
         * <return>Boolean value that indicates the potential existence of the file.</return>*
         */
        public static bool Find_File(string file_path)
        {
            if (File.Exists(file_path))
                return true;
            return false;
        }
    }
}
