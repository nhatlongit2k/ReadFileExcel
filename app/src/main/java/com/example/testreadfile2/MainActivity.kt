package com.example.testreadfile2

import android.content.Intent
import android.content.pm.PackageManager
import android.content.res.Configuration
import android.os.Bundle
import android.util.Log
import android.widget.Button
import android.widget.TextView
import androidx.appcompat.app.AppCompatActivity
import com.jaiselrahman.filepicker.activity.FilePickerActivity
import com.jaiselrahman.filepicker.config.Configurations
import com.jaiselrahman.filepicker.model.MediaFile
import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.poifs.filesystem.POIFSFileSystem
import java.io.File
import java.io.FileInputStream
import java.io.InputStream

class MainActivity : AppCompatActivity() {

    lateinit var btSelectFile: Button

    var strPath: String =""

    var TAG = "main"
    private var textView: TextView? = null
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)
        textView = findViewById(R.id.textview)
//        readExcelFileFromAssets()

        btSelectFile = findViewById(R.id.bt_select_file)

        btSelectFile.setOnClickListener {
            val intent: Intent = Intent(this, FilePickerActivity::class.java)
            intent.putExtra(
                FilePickerActivity.CONFIGS,
                Configurations.Builder().setCheckPermission(true).setShowFiles(true)
                    .setShowImages(false).setShowImages(false).setMaxSelection(1).setSuffixes("xls")
                    .setSkipZeroSizeFiles(true).build()
            )
            startActivityForResult(intent, 102)
        }
    }

//    override fun onRequestPermissionsResult(
//        requestCode: Int,
//        permissions: Array<out String>,
//        grantResults: IntArray
//    ) {
//        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
//        if((grantResults.size > 0) && (grantResults[0] == PackageManager.PERMISSION_GRANTED)){
//            if(requestCode == 1)
//        }
//    }

    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        super.onActivityResult(requestCode, resultCode, data)
        if(resultCode == RESULT_OK && data != null){
            var mediaFiles: ArrayList<MediaFile> = data.getParcelableArrayListExtra(FilePickerActivity.MEDIA_FILES)!!

            var path: String = mediaFiles.get(0).path

            when(requestCode){
                102 -> {
                    Log.d("TAG", "filePath: $path")
                    strPath = path
                    readExcelFileFromAssets(strPath)
                }
            }
        }
    }

    fun readExcelFileFromAssets() {
        try {
            val myInput: InputStream
            // initialize asset manager
            val assetManager = assets
            //  open excel sheet
            myInput = assetManager.open("myexcelsheet.xls")
            // Create a POI File System object
            val myFileSystem = POIFSFileSystem(myInput)
            // Create a workbook using the File System
            val myWorkBook = HSSFWorkbook(myFileSystem)
            // Get the first sheet from workbook
            val mySheet = myWorkBook.getSheetAt(0)
            // We now need something to iterate through the cells.
            val rowIter = mySheet.rowIterator()
            var rowno = 0
            textView!!.append("\n")
            while (rowIter.hasNext()) {
                Log.e(TAG, " row no $rowno")
                val myRow = rowIter.next() as HSSFRow
                if (rowno != 0) {
                    val cellIter = myRow.cellIterator()
                    var colno = 0
                    var sno = ""
                    var date = ""
                    var det = ""
                    while (cellIter.hasNext()) {
                        val myCell = cellIter.next() as HSSFCell
                        if (colno == 0) {
                            sno = myCell.toString()
                        } else if (colno == 1) {
                            date = myCell.toString()
                        } else if (colno == 2) {
                            det = myCell.toString()
                        }
                        colno++
                        Log.e(TAG, " Index :" + myCell.columnIndex + " -- " + myCell.toString())
                    }
                    textView!!.append("$sno -- $date  -- $det\n")
                }
                rowno++
            }
        } catch (e: Exception) {
            Log.e(TAG, "error $e")
        }
    }

    fun readExcelFileFromAssets(path: String) {
        try {
            var myInput: InputStream
            // initialize asset manager
//            val assetManager = assets
            //  open excel sheet
//            myInput = assetManager.open("myexcelsheet.xls")

            val file: File = File(path)
            myInput = FileInputStream(file)
            // Create a POI File System object
            val myFileSystem = POIFSFileSystem(myInput)
            // Create a workbook using the File System
            val myWorkBook = HSSFWorkbook(myFileSystem)
            // Get the first sheet from workbook
            val mySheet = myWorkBook.getSheetAt(0)
            // We now need something to iterate through the cells.
            val rowIter = mySheet.rowIterator()
            var rowno = 0
            textView!!.append("\n")
            while (rowIter.hasNext()) {
                Log.e(TAG, " row no $rowno")
                val myRow = rowIter.next() as HSSFRow
                if (rowno != 0) {
                    val cellIter = myRow.cellIterator()
                    var colno = 0
                    var sno = ""
                    var date = ""
                    var det = ""
                    while (cellIter.hasNext()) {
                        val myCell = cellIter.next() as HSSFCell
                        if (colno == 0) {
                            sno = myCell.toString()
                        } else if (colno == 1) {
                            date = myCell.toString()
                        } else if (colno == 2) {
                            det = myCell.toString()
                        }
                        colno++
                        Log.e(TAG, " Index :" + myCell.columnIndex + " -- " + myCell.toString())
                    }
                    textView!!.append("$sno -- $date  -- $det\n")
                }
                rowno++
            }
        } catch (e: Exception) {
            Log.e(TAG, "error $e")
        }
    }
}