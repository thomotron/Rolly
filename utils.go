package main

import (
	"log"
)

// Logs a fatal error if err is not nil
func fatalErr(err error, message string) {
	if err != nil {
		if len(message) > 0 {
			log.Fatalf("%v: %v", message, err.Error())
		} else {
			log.Fatal(err.Error())
		}
	}
}
