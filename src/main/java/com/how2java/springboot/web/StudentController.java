package com.how2java.springboot.web;

import com.how2java.springboot.entity.Student;
import com.how2java.springboot.service.StudentService;
import com.how2java.springboot.util.ExcelUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.util.List;

@Controller
public class StudentController {
    @Autowired
    StudentService studentService;

    @RequestMapping("excel")
    public String excel() {
        return "excel";
    }

    @PostMapping("/importExcel")
    public String importData(HttpServletRequest req, @RequestParam("file") MultipartFile file){
//        String filename = file.getOriginalFilename();
        String filename = "C:/Users/XinjianLi/Desktop/student1.xlsx";
        List<Object[]> list = ExcelUtil.importExcel(filename);
        for(int i= 0;i<list.size();i++){
            Student student  = new Student();
            student.setId((Integer)list.get(i)[0]);
            student.setName((String) list.get(i)[1]);
            student.setAge((Integer) list.get(i)[2]);
//            System.out.println(student.toString());
            studentService.importData(student);
        }

        /*将文件上传到目的目录*/
        try {
            String fileName1 = System.currentTimeMillis()+file.getOriginalFilename();
            String destFileName=req.getServletContext().getRealPath("")+"uploaded"+ File.separator+fileName1;

            File destFile = new File(destFileName);
            destFile.getParentFile().mkdirs();
            file.transferTo(destFile);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
            return "上传失败," + e.getMessage();
        } catch (IOException e) {
            e.printStackTrace();
            return "上传失败," + e.getMessage();
        }
        return "success";
    }

    @PostMapping("/exportExcel")
    public void exportData(HttpServletResponse response) throws UnsupportedEncodingException {
        List<Student> studentList =studentService.findStudentList();
        String[] tableHeader = {"id","name","age"};
        ExcelUtil.exportExcel(response,studentList,tableHeader);
    }

}
