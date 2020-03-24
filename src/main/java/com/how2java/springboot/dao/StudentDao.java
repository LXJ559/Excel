package com.how2java.springboot.dao;

import com.how2java.springboot.entity.Student;
import org.springframework.data.jpa.repository.JpaRepository;


public interface StudentDao extends JpaRepository<Student,Integer> {
}
