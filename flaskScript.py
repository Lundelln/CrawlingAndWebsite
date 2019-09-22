
import os
from flask import Flask, render_template, request, send_file
import dealCrawlerWJ
import dealCrawlerES

app = Flask(__name__)

APP_ROOT = os.path.dirname(os.path.abspath(__file__))

@app.route("/", methods=['GET', 'POST'])
def home():
    return render_template("main.html")

@app.route('/wj/', methods=['GET', 'POST'])
def WJ():
	try:
		excelFile = dealCrawlerWJ.comparePrices()
		return render_template("wj.html", excelFile = excelFile)
	except Exception as e:
		return str(e)

@app.route('/download_wj/', methods=['GET', 'POST'])
def downloadWJ():
	path = "wjPrices.xlsx"
	return send_file(path, as_attachment=True)

@app.route('/es/', methods=['GET', 'POST'])
def ES():
	try:
		excelFile = dealCrawlerES.comparePrices()
		return render_template("es.html", excelFile = excelFile)
	except Exception as e:
		return str(e)


@app.route('/download_es/', methods=['GET', 'POST'])
def downloadES():
	path = "elektroskandia.xlsx"
	return send_file(path, as_attachment=True)

@app.route('/upload/', methods=['POST'])
def upload():
	target = APP_ROOT
	print(target)

	for file in request.files.getlist("file"):
		print(file)
		filename = file.filename
		destination = "/".join([target, filename])
		print(destination)
		file.save(destination)

	return render_template("main.html")


@app.errorhandler(404)
def page_not_found(e):
	return render_template("404.html")

if __name__ == '__main__':
	app.run(debug=True)
