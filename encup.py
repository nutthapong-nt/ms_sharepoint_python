import rsa

#complete user/pass here run save encrypt to config don't forget delete username and password after done
username=""
password=""

key=rsa.PublicKey("*secret*")

enusername=rsa.encrypt(username.encode(),key)
enpassword=rsa.encrypt(password.encode(),key)

f=open("userconfig.py","wt")
print(f'enusername={enusername}',file=f)
print(f'enpassword={enpassword}',file=f)
#f.writelines(f'"enusername":{enusername}')
#f.writelines(f'"enpassword":{enpassword}')
f.close()