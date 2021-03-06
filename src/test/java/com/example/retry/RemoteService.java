package com.example.retry;

/**
 * Created by SDD on 2017/4/24.
 */
import org.springframework.remoting.RemoteAccessException;
import org.springframework.retry.annotation.Backoff;
import org.springframework.retry.annotation.Recover;
import org.springframework.retry.annotation.Retryable;
import org.springframework.stereotype.Service;

@Service
public class RemoteService {
	@Retryable(value= {RemoteAccessException.class},maxAttempts = 3,backoff = @Backoff(delay = 5000l,multiplier = 1))
	public void call() throws Exception {
		System.out.println("do something...");
		throw new RemoteAccessException("RPC调用异常");
	}
	@Recover
	public void recover(RemoteAccessException e) {
		System.out.println(e.getMessage());
	}
}
