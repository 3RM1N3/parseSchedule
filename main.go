package main

import "fmt"

func main() {
	err := ParseFile()
	if err != nil {
		fmt.Println(err)
		return
	}
}
