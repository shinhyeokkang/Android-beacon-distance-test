package org.altbeacon.beaconreference;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Date;
import java.util.LinkedList;

import android.app.Activity;

import android.os.Bundle;
import android.os.RemoteException;
import android.util.Log;
import android.widget.EditText;
import android.widget.Toast;

import org.altbeacon.beacon.AltBeacon;
import org.altbeacon.beacon.Beacon;
import org.altbeacon.beacon.BeaconConsumer;
import org.altbeacon.beacon.BeaconManager;
import org.altbeacon.beacon.BeaconParser;
import org.altbeacon.beacon.RangeNotifier;
import org.altbeacon.beacon.Region;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RangingActivity extends Activity implements BeaconConsumer {
    protected static final String TAG = "RangingActivity";
    private BeaconManager beaconManager = BeaconManager.getInstanceForApplication(this);
    private int[] rssiList = new int[1000000];
    private int Tx;
    private int count=0;
    //private int num=1;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_ranging);
    }


    // xls 형식으로 결과를 저장하는 메소드
    private void saveExcel() throws IOException { // rssiList를 받아 엑셀파일로 작성
        Workbook workbook = new HSSFWorkbook();

        Sheet sheet = workbook.createSheet(); // 새로운 시트 생성

        Row row = sheet.createRow(0); // 새로운 행 생성
        Cell cell;

        cell = row.createCell(0); // 1번 셀 생성
        cell.setCellValue("TxPower"); // 1번 셀 값 입력

        cell = row.createCell(1); // 2번 셀 생성
        cell.setCellValue("RSSI"); // 2번 셀 값 입력

        for(int i = 0; i < 50; i++){ // 데이터 엑셀에 입력
            row = sheet.createRow(i+1);
            cell = row.createCell(0);
            cell.setCellValue(Tx);
            cell = row.createCell(1);
            cell.setCellValue(rssiList[i]);
        }
        //파일명을 현재시각으로 설정하는 코드
        long mNow = System.currentTimeMillis();
        Date mReDate = new Date(mNow);
        SimpleDateFormat mFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String formatDate = mFormat.format(mReDate);

        File xlsFile = new File(getExternalFilesDir(null),formatDate+"_test.xls");
        //num++; // 파일 생성시 마다 숫자 증가시켜서 중복 회피
        try{
            FileOutputStream os = new FileOutputStream(xlsFile);
            workbook.write(os); // 외부 저장소에 엑셀 파일 생성
        }catch (IOException e){
            e.printStackTrace();
        }
        Toast.makeText(getApplicationContext(),xlsFile.getAbsolutePath()+"에 저장되었습니다",Toast.LENGTH_SHORT).show();
    }

    // xlsx 형식으로 파일을 저장하는 메소드
    private void saveExcelX() throws IOException { // rssiList를 받아 엑셀파일로 작성

        String sheetName = "Sheet1";//name of sheet

        XSSFWorkbook workbook = new XSSFWorkbook(); // xlsx 파일 형식으로 변경

        XSSFSheet sheet = workbook.createSheet(sheetName); // 새로운 시트 생성

        XSSFRow row = sheet.createRow(0); // 새로운 행 생성
        XSSFCell cell;

        cell = row.createCell(0); // 1번 셀 생성
        cell.setCellValue("TxPower"); // 1번 셀 값 입력

        cell = row.createCell(1); // 2번 셀 생성
        cell.setCellValue("RSSI"); // 2번 셀 값 입력

        for(int i = 0; i < 50; i++){ // 데이터 엑셀에 입력
            row = sheet.createRow(i+1);
            cell = row.createCell(0);
            cell.setCellValue(Tx);
            cell = row.createCell(1);
            cell.setCellValue(rssiList[i]);
        }
        //파일명을 현재시각으로 설정하는 코드
        long mNow = System.currentTimeMillis();
        Date mReDate = new Date(mNow);
        SimpleDateFormat mFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String formatDate = mFormat.format(mReDate);

        File xlsxFile = new File(getExternalFilesDir(null),formatDate+"_test.xlsx");
        //num++; // 파일 생성시 마다 숫자 증가시켜서 중복 회피
        try{
            FileOutputStream os = new FileOutputStream(xlsxFile);
            workbook.write(os); // 외부 저장소에 엑셀 파일 생성
        }catch (IOException e){
            e.printStackTrace();
        }
        Toast.makeText(getApplicationContext(),xlsxFile.getAbsolutePath()+"에 저장되었습니다",Toast.LENGTH_SHORT).show();
    }


    @Override 
    protected void onDestroy() {
        super.onDestroy();
    }

    @Override 
    protected void onPause() {
        super.onPause();
        beaconManager.unbind(this);
    }

    @Override 
    protected void onResume() {
        super.onResume();
        beaconManager.bind(this);
    }

    @Override
    public void onBeaconServiceConnect() {

        RangeNotifier rangeNotifier = new RangeNotifier() {
           @Override
           public void didRangeBeaconsInRegion(Collection<Beacon> beacons, Region region) {
              if (beacons.size() > 0) {
                  Log.d(TAG, "didRangeBeaconsInRegion called with beacon count:  "+beacons.size());
                  Beacon firstBeacon = beacons.iterator().next();
                  Tx = firstBeacon.getTxPower();
                  logToDisplay("The first beacon " + firstBeacon.toString() + " is about " + firstBeacon.getDistance() + " meters away. \n" +"Rssi: " + firstBeacon.getRssi() + " TxPower: " + firstBeacon.getTxPower());
                  rssiList[count] = firstBeacon.getRssi();
                  if(count==50){
                      try {
                          saveExcelX();
                      } catch (IOException e) {
                          e.printStackTrace();
                      }
                      count=0; // 초기화
                      logToDisplay("List has filled!!!!!!!!!!!!"); // 완성 문구 출력
                  }
                  count++;
              }
           }

        };
        try {
            beaconManager.startRangingBeaconsInRegion(new Region("myRangingUniqueId", null, null, null));
            beaconManager.addRangeNotifier(rangeNotifier);
            beaconManager.startRangingBeaconsInRegion(new Region("myRangingUniqueId", null, null, null));
            beaconManager.addRangeNotifier(rangeNotifier);
        } catch (RemoteException e) {   }
    }

    private void logToDisplay(final String line) {
        runOnUiThread(new Runnable() {
            public void run() {
                EditText editText = (EditText)RangingActivity.this.findViewById(R.id.rangingText);
                editText.append(line+"\n");
            }
        });
    }
}
