from twocaptcha import TwoCaptcha
import time



solver = TwoCaptcha('8f63da7191fe11e63148c3d8b28c71f2')

id = solver.send(file='captcha-bg.jpg')
time.sleep(20)

code = solver.get_result(id)
print(code)