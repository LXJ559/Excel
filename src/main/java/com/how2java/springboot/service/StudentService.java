package com.how2java.springboot.service;

import com.how2java.springboot.dao.StudentDao;
import com.how2java.springboot.entity.Student;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
public class StudentService {
    @Autowired
    StudentDao studentDao;

    public void importData(Student student){
        studentDao.save(student);
    }

    public List<Student> findStudentList(){ return studentDao.findAll();}
}
