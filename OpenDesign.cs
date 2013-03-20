using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using Teigha.DatabaseServices;
using Teigha.Runtime;

namespace Teigha
{
    public class OpenDesign
    {
        Teigha.Runtime.Services dd;
        //Graphics graphics;
        //Teigha.GraphicsSystem.LayoutHelperDevice helperDevice;
        //Database database = null;
        public OpenDesign()
        {
            dd = new Teigha.Runtime.Services();
        }
        ~OpenDesign()
        {
            if (dd != null)
                dd.Dispose();
        }
        public Bitmap Thumb(string FullDwgPath)
        {
            Bitmap resBMP = null;
            try
            {
                Database db = new Database(false, false);
                try
                {
                    db.ReadDwgFile(FullDwgPath, FileOpenMode.OpenForReadAndAllShare, false, "");
                    Bitmap bmp = db.ThumbnailBitmap;
                    if (bmp != null)
                    {
                        resBMP = bmp.Clone() as Bitmap;
                    }
                }
                catch (Teigha.Runtime.Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("In Reading bitmap \n" + ex.Message);
                }
            }
            catch (Teigha.Runtime.Exception rex)
            {
                System.Windows.Forms.MessageBox.Show("In creating database \n" + rex.Message);
            }

            return resBMP;
        }


    }
}
