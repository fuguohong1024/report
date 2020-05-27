package main

import (
	"bufio"
	"database/sql"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	//调用初始化函数
	_ "github.com/go-sql-driver/mysql"
	"github.com/jordan-wright/email"
	"github.com/robfig/cron"
	"io"
	"log"
	"net/smtp"
	"os"
	"strconv"
	"strings"
	"test/conf"
	"time"
)

//日志配置
func init(){
	log.SetPrefix("[report]")
	log.SetFlags(log.LstdFlags | log.Lshortfile)
}


//发送邮件函数

func sendMail(){
	t := time.Now().Format("2006-01-02")
	//调用NewEmail
	e := email.NewEmail()
	//toml读取配置
	mailconf := conf.Conf.Mail
	name := mailconf.Name
	to := mailconf.To
	//附件名
	f := t+name+".xlsx"
	a,_ := e.AttachFile(f)
	e.From = "fgh <17713603104m@sina.cn>"
	e.To = to
	//e.Bcc = []string{"test_bcc@example.com"}
	//e.Cc = []string{"test_cc@example.com"}
	//创建切片slice
	slice := make([]*email.Attachment,0)
	//切片添加Attachment值
	slice = append(slice, a)
	e.Subject = name
	//e.Text = []byte("Text Body is, of course, supported!")
	e.HTML = []byte("<h1>附件为"+name+"数据，请查收</h1>")
	e.Attachments = slice
	err := e.Send("smtp.sina.com:587", smtp.PlainAuth("", "17713603104m@sina.cn", "授权码", "smtp.sina.com"))
	if err != nil {
		log.Println(err)
	}else {
		log.Println("send mail suceessful!")
	}
}


func sqltoexcel() {
	start := time.Now()
	//所有标志都声明完成以后，调用 flag.parse() 来执行命令行解析
	flag.Parse()
	//tmol读取配置文件
	if err := conf.Init(); err != nil {
		log.Println("conf.Init() err:%+v", err)
	}
	mysqlconf := conf.Conf.Mysql
	mailconf := conf.Conf.Mail
	dsn := mysqlconf.Dsn
	query := mysqlconf.Query
	name := mailconf.Name
	//spec := mailconf.Spec
	//键值对取配置参数
	/*config := InitConfig(*Path)
	dsn := config["dsn"]
	query := config["query"]
	name := config["name"]*/
	// 打开数据连接
	db, err := sql.Open("mysql", dsn)
	//封装的错误处理函数调用
	checkErr(err)
	//延迟（defer）关闭数据库连接,待函数执行完
	defer db.Close()

	//获取查询值
	resultPointer, columnsPointer := sqlFetch(db, query)
	//传入excel函数处理写入
	excel(resultPointer, columnsPointer,name)
	end := time.Now()
	//打印日志
	log.Println("sql:",query)
	log.Println("query time : ", timeFriendly(end.Sub(start).Seconds()))
	sendMail()
}



func main() {
	flag.Parse()
	//tmol读取配置文件
	if err := conf.Init(); err != nil {
		log.Println("conf.Init() err:%+v", err)
	}
	mailconf := conf.Conf.Mail
	spec := mailconf.Spec
	log.Println("定时：",spec)
	// 新建一个定时任务对象
	c := cron.New()
	c.AddFunc(spec,sqltoexcel)
	c.Run()
	select {

	}
}

func cronlog(){
	t := time.Now()
	log.Println(t)
}

//获取字段名和值
func sqlFetch(db *sql.DB, query string) (*[]map[string]string, *[]string) {

	//执行查询
	rows, err := db.Query(query)
	checkErr(err)
	//获取字段名
		columns, err := rows.Columns()
	checkErr(err)
	//创建切片存入数据
	values := make([]sql.RawBytes, len(columns))
	//rows.Scan 需要传入[]interface{}类型的参
	scanArgs := make([]interface{}, len(values))
	//遍历值传入scanArgs[]interface{}
	for i := range values {
		scanArgs[i] = &values[i]
	}
	//初始化result键值对
	result := make([]map[string]string, 0)
	//获取行
	for rows.Next() {
		//获取数据
		err = rows.Scan(scanArgs...)
		checkErr(err)
		//处理数据,将每行打印为string型
		var value string
		vmap := make(map[string]string, len(scanArgs))
		for i, col := range values {
			//空值赋为NULL
			if col == nil {
				value = "NULL"
			} else {
				value = string(col)
			}
			vmap[columns[i]] = value
			//值传入vmap
		}
		//传入result
		result = append(result, vmap)
	}
	//错误处理
	if err = rows.Err(); err != nil {
		panic(err.Error())
	}
	return &result, &columns

}

func excel(resultPointer *[]map[string]string, columnsPointer *[]string,name string) {
	//创建excel
	xlsx := excelize.NewFile()

	//设置单元格的值
	result := *resultPointer
	columns := *columnsPointer


	//字段名categories := map[string]string{"A1": "Small", "B1": "Apple", "C1": "Orange", "D1": "Pear"}
	//值values := map[string]int{"B2": 2, "C2": 3, "D2": 3, "B3": 5, "C3": 2, "D3": 4, "B4": 6, "C4": 7, "D4": 8}

	categories := make(map[string]string, 0)

	//写入字段名
	for k, v := range columns {
		key := precessCategories(k)
		categories[key+"1"] = v
	}
	//fmt.Println(categories)
	for k, v := range categories {
		xlsx.SetCellValue("Sheet1", k, v)
	}

	values := make(map[string]string, 0)

	for k1, v1 := range result {

		//fmt.Println(v1)
		c := 0
		for k2, v2 := range v1 {

			i := getArrKey(columns, k2)
			key := precessCategories(i) + strconv.Itoa(k1+2)
			values[key] = v2
			//fmt.Println(key)

			c++
		}

	}
	//写入值)
	for k, v := range values {
		xlsx.SetCellValue("Sheet1", k, v)
	}

	// Set active sheet of the workbook.
	xlsx.SetActiveSheet(2)
	// Save xlsx file by the given path.
	t := time.Now().Format("2006-01-02")
	err := xlsx.SaveAs(t+name+".xlsx")
	if err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

//map is not ordered
func getArrKey(arr []string, value string) int {
	for k, v := range arr {
		if v == value {
			return k
		}
	}
	return -1
}

//excel行设置
func precessCategories(k int) string {
	az := "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	if k < 26 {
		return string(az[k])
	} else {
		k1 := int((k + 1) / 26)
		k2 := (k + 1) % 26
		return string(az[k1]) + string(az[k2])
	}
}

func checkErr(err error) {
	if err != nil {
		log.Println(err)
	}
}

//格式化时间
func timeFriendly(second float64) string {

	if second < 1 {
		return strconv.Itoa(int(second*1000)) + "毫秒"
	} else if second < 60 {
		return strconv.Itoa(int(second)) + "秒" + timeFriendly(second-float64(int(second)))
	} else if second >= 60 && second < 3600 {
		return strconv.Itoa(int(second/60)) + "分" + timeFriendly(second-float64(int(second/60)*60))
	} else if second >= 3600 && second < 3600*24 {
		return strconv.Itoa(int(second/3600)) + "小时" + timeFriendly(second-float64(int(second/3600)*3600))
	} else if second > 3600*24 {
		return strconv.Itoa(int(second/(3600*24))) + "天" + timeFriendly(second-float64(int(second/(3600*24))*(3600*24)))
	}
	return ""
}


//读取key=value类型的配置文件
func InitConfig(path string) map[string]string {
	config := make(map[string]string)

	f, err := os.Open(path)
	defer f.Close()
	if err != nil {
		panic(err)
	}

	r := bufio.NewReader(f)
	for {
		b, _, err := r.ReadLine()
		if err != nil {
			if err == io.EOF {
				break
			}
			panic(err)
		}
		s := strings.TrimSpace(string(b))
		index := strings.Index(s, "=")
		if index < 0 {
			continue
		}
		key := strings.TrimSpace(s[:index])
		if len(key) == 0 {
			continue
		}
		value := strings.TrimSpace(s[index+1:])
		if len(value) == 0 {
			continue
		}
		config[key] = value
	}
	return config
}
