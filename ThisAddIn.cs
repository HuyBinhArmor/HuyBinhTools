using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace HuyBinhTools
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public Excel.Workbook GetActiveWB() => (Excel.Workbook)Application.ActiveWorkbook;
        public Excel.Worksheet GetAciveWS() => (Excel.Worksheet)Application.ActiveSheet;
        public Excel.Range GetSelection() => (Excel.Range)Application.Selection;
        public Excel.Range GetLastCellFilled() => (Excel.Range)Globals.ThisAddIn.GetAciveWS().Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell,Type.Missing);

       

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    } // Hàm hệ thống

    public class Sotienbangchu
    {
        private static string [] m09Text = new string[10] { " không", " một", " hai", " ba", " bốn", " năm", " sáu", " bảy", " tám", " chín" };
        private static string [] mHauto = new string[6] { "", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ" };
        private static string BobasoText(int _BoBaSo)
        {
            int _SoHangTram, _SoHangChuc, _SoHangDvi;
            string _mString = "";
            _SoHangTram = (int)(_BoBaSo / 100);
            _SoHangChuc = (int)((_BoBaSo % 100) / 10);
            _SoHangDvi = _BoBaSo % 10;

            if ((_SoHangTram == 0) && (_SoHangChuc == 0) && (_SoHangDvi == 0)) return "";
            // Xét số hàng trăm
            _mString += m09Text[_SoHangTram] + " trăm";
            // Xét số hàng chục
            if ((_SoHangChuc == 0) && (_SoHangDvi == 0)) return _mString;
            if ((_SoHangChuc == 0) && (_SoHangDvi > 0)) _mString += " linh";
            if (_SoHangChuc == 1) _mString += " mười";
            if (_SoHangChuc > 1) _mString += m09Text[_SoHangChuc] + " mươi";
            // Xét số hàng đơn vị
            switch (_SoHangDvi)
            {
                case 1:
                    if (_SoHangChuc > 1)
                    {
                        _mString += " mốt";
                    }
                    else
                    {
                        _mString += " một";
                    }
                    break;
                case 5:
                    if (_SoHangChuc == 0)
                    {
                        _mString += " năm";
                    }
                    else
                    {
                        _mString += " lăm";
                    }
                    break;
                default:
                    if (_SoHangDvi != 0)
                    {
                        _mString += m09Text[_SoHangDvi];
                    }
                    break;
            }

            return _mString;
        } // Hàm đọc bộ 3 chữ số

        public string DocSoTienBangChu(double _SoTien)
        {
            int _SoLop, i;
            string _Dau = "";
            string _mString = "", _Boba = "";
            int[] Lop = new int[6];
            if (_SoTien == 0) return "Không";
            if (_SoTien < 0)
            {
                _SoTien = -_SoTien;
                _Dau = "Âm ";
            }

            //Kiểm tra số quá lớn
            if (_SoTien > 9000000000000000)
            {
                return "";
            }
            Lop[5] = (int)(_SoTien / Math.Pow(10, 15));
            _SoTien = (long)(_SoTien % Math.Pow(10, 15));

            Lop[4] = (int)(_SoTien / Math.Pow(10, 12));
            _SoTien = (long)(_SoTien % Math.Pow(10, 12));

            Lop[3] = (int)(_SoTien / Math.Pow(10, 9));
            _SoTien = (long)(_SoTien % Math.Pow(10, 9));

            Lop[2] = (int)(_SoTien / Math.Pow(10, 6));
            _SoTien = (long)(_SoTien % Math.Pow(10, 6));

            Lop[1] = (int)(_SoTien / Math.Pow(10, 3));
            _SoTien = (long)(_SoTien % Math.Pow(10, 3));

            Lop[0] = (int)(_SoTien / Math.Pow(10, 0));




            if (Lop[5] > 0)
            {
                _SoLop = 5;
            }
            else if (Lop[4] > 0)
            {
                _SoLop = 4;
            }
            else if (Lop[3] > 0)
            {
                _SoLop = 3;
            }
            else if (Lop[2] > 0)
            {
                _SoLop = 2;
            }
            else if (Lop[1] > 0)
            {
                _SoLop = 1;
            }
            else
            {
                _SoLop = 0;
            }
            for (i = _SoLop; i >= 0; i--)
            {
                _Boba = BobasoText(Lop[i]);
                _mString += _Boba;

                if (Lop[i] != 0) _mString += mHauto[i];
                if ((i > 0) && (!string.IsNullOrEmpty(_Boba))) _mString += ",";
            }


            if (_mString.Substring(0, 16) == " không trăm linh")
            {
                _mString = _mString.Substring(16, _mString.Length - 16);
            }
            else if (_mString.Substring(0, 11) == " không trăm")
            {
                _mString = _mString.Substring(11, _mString.Length - 11);
            }

            _mString = _Dau + _mString.Trim() + " đồng./."; // Thêm dấu âm và đơn vị tính
            return _mString.Substring(0, 1).ToUpper() + _mString.Substring(1); // Viết hoa chữ đầu
        }


    } // Hàm đọc số tiền bằng chữ


}
