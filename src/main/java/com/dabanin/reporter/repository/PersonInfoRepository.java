package com.dabanin.reporter.repository;

import com.dabanin.reporter.entity.PersonInfo;
import org.springframework.data.repository.CrudRepository;

public interface PersonInfoRepository extends CrudRepository<PersonInfo, Long> {
}
