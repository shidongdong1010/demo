package com.example.retry;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.retry.annotation.EnableRetry;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

/**
 * Created by SDD on 2017/4/24.
 */
@RunWith(SpringJUnit4ClassRunner.class)
@SpringBootTest
@EnableRetry
public class RetryTest {

	@Autowired
	private RemoteService remoteService;

	@Test
	public void test() throws Exception {
		remoteService.call();
	}
}
