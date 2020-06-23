from flask import Flask,render_template,url_for, redirect,request
import Gmailmain
import os

app = Flask(__name__)


@app.route('/success/<count>/') 
def success(count): 
        return render_template('sucess.html',count1 = count)

@app.route('/')
def Home_page():
        return render_template('Home1.html')

@app.route('/login',methods=['POST','GET'])
def Second_page():
        if request.method=="POST":
                drivename = request.form['drive']
                foldername = request.form['foldername']
                idfield = request.form['idname']
                # sheet = request.form['sheetname'] 
                subject = request.form['Query']
                date = request.form['date']
                file1 = request.files['db']
                string = drivename+foldername+'\\'
                if not os.path.exists(string):
                        os.makedirs(string)
                file1.save(os.path.join(string, file1.filename))
                
                
                db1 =  file1.filename
                
                print(string)
                dirstring = string.replace(os.sep,'/')

                COUNT = Gmailmain.main(subject,date,db1,dirstring,idfield)
                return redirect(url_for('success',count = COUNT))

if __name__ == '__main__':
    app.run(debug=True)
