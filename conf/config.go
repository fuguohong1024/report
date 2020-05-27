package conf

import (
	"flag"
	"github.com/BurntSushi/toml"
)

var (
	confPath string
	//Conf 全局变量
	Conf = &Config{}
)

// Config .
type Config struct {
	Title string
	Mysql Mysql
	Mail Mail
}

type Mysql struct {
	Dsn string
	Query string
}

type Mail struct {
	Spec string
	Name string
	To []string
}



func init() {
	flag.StringVar(&confPath, "conf", "./conf.toml", "-conf path")
}


//初始化配置
func Init() (err error) {
	_, err = toml.DecodeFile(confPath, &Conf)
	return

}
