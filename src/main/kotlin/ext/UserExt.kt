package ext

import User
import XrHelloUtil

fun User.xrPrint() {
    println("我是 User 的扩展函数")
}

fun XrHelloUtil.xrHello(){
    println("我是 XrHelloUtil 的扩展函数 ")

}