package com.example.asyn;


import org.junit.Test;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.concurrent.CountDownLatch;
import java.util.concurrent.TimeUnit;

/**
 * Created by SDD on 2017/3/7.
 */
@SpringBootTest
public class AsynMethod {

	private void asyMethod( String param,RequestCallback callback){
		heavyWork();
		callback.callback(param);
		System.out.println("执行异步方法结束");
	}

	private void heavyWork(){
		while (true){
		}
		/*try {
			TimeUnit.SECONDS.sleep(30);
		} catch (InterruptedException e) {
		}*/
	}

	@Test
	public void getHeavyCallbackResult2(){
		System.out.println("调用异步方法开始");
		final String[] a = {null};
		final CountDownLatch latch = new CountDownLatch(1);
		new Thread(new Runnable() {
			@Override
			public void run() {
				asyMethod("getHeavyCallbackResult2",new RequestCallback() {
					@Override
					public void callback(String result) {
						a[0] = result;
						latch.countDown();
					}
				});
			}
		}).start();

		try {
			latch.await();
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		System.out.println("调用异步方法结束: " + a[0]);
	}

	interface RequestCallback{
		void callback(String param);
	}
}
