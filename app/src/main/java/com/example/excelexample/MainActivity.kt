package com.example.excelexample

import android.Manifest
import android.content.pm.PackageManager
import androidx.appcompat.app.AppCompatActivity
import android.os.Bundle
import android.os.Environment
import android.text.TextUtils
import android.util.Log
import android.widget.Toast
import androidx.core.app.ActivityCompat
import com.aspose.cells.BorderType
import com.aspose.cells.CellBorderType
import com.aspose.cells.Color
import com.aspose.cells.TextAlignmentType
import com.aspose.cells.Workbook
import com.example.excelexample.databinding.ActivityMainBinding
import java.io.File
import java.io.FileNotFoundException
import java.io.FileOutputStream
import java.io.IOException
import java.io.OutputStream
import java.io.OutputStreamWriter

class MainActivity : AppCompatActivity() {
    lateinit var binding: ActivityMainBinding
    var TAG : String = "MainActivity"
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        binding = ActivityMainBinding.inflate(layoutInflater)
        setContentView(binding.root)

        binding.cetakbtn.setOnClickListener {
            if (ActivityCompat.checkSelfPermission(applicationContext,Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED){
                ActivityCompat.requestPermissions(this, arrayOf(Manifest.permission.WRITE_EXTERNAL_STORAGE),1)
                return@setOnClickListener
            }else if (TextUtils.isEmpty(binding.edtNamaLengkap.text.toString())){
                Toast.makeText(applicationContext,"Nama Lengkap Tidak Boleh Kosong",Toast.LENGTH_LONG).show()
            }else if (TextUtils.isEmpty(binding.edtUsiaAnda.text.toString())){
               Toast.makeText(applicationContext,"Usia Tidak Boleh Kosong",Toast.LENGTH_LONG).show()
            }else if (TextUtils.isEmpty(binding.edtPendidikanAnda.text.toString())){
                Toast.makeText(applicationContext,"Pendidikan Tidak Boleh Kosong",Toast.LENGTH_LONG).show()
            }else{
                var workbook : Workbook = Workbook()
                var sheetIndex : Int = workbook.worksheets.add()
                Log.d(TAG, "onCreate sheetIndex: ${sheetIndex}")
                var worksheet = workbook.worksheets.get(sheetIndex)
                Log.d(TAG, "onCreate worksheet: ${worksheet}")

                var cells = worksheet.cells

                //CellNama
                cells.setColumnWidth(0, 15.0)
                cells.setRowHeight(0,15.0)
                cells.get(0,0).value = "Nama Lengkap"
                var style = cells.get(0,0).style
                style.font.isBold = true
                style.horizontalAlignment = TextAlignmentType.CENTER
                style.setBorder(BorderType.TOP_BORDER,CellBorderType.THICK,Color.getBlack())
                style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,Color.getBlack())
                style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK,Color.getBlack())
                style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK, Color.getBlack())
                cells.get(0,0).setStyle(style)

                //isian nama
                cells.setColumnWidth(1,15.0)
                cells.setRowHeight(1,15.0)
                cells.get(1,0).value = binding.edtNamaLengkap.text.toString()
                var style1 = cells.get(1,1).style
                style1.horizontalAlignment = TextAlignmentType.CENTER
                style1.setBorder(BorderType.TOP_BORDER,CellBorderType.THICK,Color.getBlack())
                style1.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK, Color.getBlack())
                style1.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK, Color.getBlack())
                style1.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK, Color.getBlack())
                cells.get(1,0).setStyle(style1)

                //Cells Usia
                cells.setColumnWidth(1,15.0)
                cells.setRowHeight(0,15.0)
                cells.get(0,1).value = "Usia"
                var styleusia = cells.get(1,1).style
                styleusia.font.isBold = true
                styleusia.horizontalAlignment = TextAlignmentType.CENTER
                styleusia.setBorder(BorderType.TOP_BORDER,CellBorderType.THICK, Color.getBlack())
                styleusia.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK, Color.getBlack())
                styleusia.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK, Color.getBlack())
                styleusia.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK, Color.getBlack())
                cells.get(0,1).setStyle(styleusia)
                //Isian Usia
                cells.get(1,1).value = binding.edtUsiaAnda.text.toString()
                var styleisianusia = cells.get(1,1).style
                styleisianusia.horizontalAlignment = TextAlignmentType.CENTER
                styleisianusia.setBorder(BorderType.TOP_BORDER,CellBorderType.THICK, Color.getBlack())
                styleisianusia.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK, Color.getBlack())
                styleisianusia.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK, Color.getBlack())
                styleisianusia.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK, Color.getBlack())
                cells.get(1,1).setStyle(styleisianusia)


                //cell pendidikan
                cells.setColumnWidth(2,15.0)
                cells.setRowHeight(0,15.0)
                cells.get(0,2).value = "Pendidikan"
                var stylepend = cells.get(0,2).style
                stylepend.font.isBold = true
                stylepend.horizontalAlignment = TextAlignmentType.CENTER
                stylepend.setBorder(BorderType.TOP_BORDER,CellBorderType.THICK, Color.getBlack())
                stylepend.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK, Color.getBlack())
                stylepend.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK, Color.getBlack())
                stylepend.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK, Color.getBlack())
                cells.get(0,2).setStyle(stylepend)

                //isian
                cells.setColumnWidth(2,15.0)
                cells.setRowHeight(1,15.0)
                cells.get(1,2).value = binding.edtPendidikanAnda.text.toString()
                var styleisianpend = cells.get(3,3).style
                styleisianpend.horizontalAlignment = TextAlignmentType.CENTER
                styleisianpend.setBorder(BorderType.TOP_BORDER,CellBorderType.THICK, Color.getBlack())
                styleisianpend.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK, Color.getBlack())
                styleisianpend.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK, Color.getBlack())
                styleisianpend.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK, Color.getBlack())
                cells.get(1,2).setStyle(styleisianpend)



                var fileName : String =  "${Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)}/Biodata.xlsx"
                Log.i(TAG, "onCreate FileName: "+fileName)

                try {

                    workbook.save(fileName)
                    Log.i(TAG, "onCreate workbook Berhasil:" )
                    Toast.makeText(applicationContext,"Cetak Excel Berhasil",Toast.LENGTH_LONG).show()
                }catch (e : FileNotFoundException){
                    Log.i(TAG, "onCreate FileNotFound: "+e.message)
                }catch (e : IOException){
                    Log.i(TAG, "onCreate IoException: "+e.message)
                }
            }

        }
    }
}