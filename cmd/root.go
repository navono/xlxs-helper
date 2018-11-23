package cmd

import (
	"fmt"

	"github.com/spf13/cobra"
)

// RootCmd root command
var RootCmd = &cobra.Command{
	Use: "hello",
	Run: func(cmd *cobra.Command, args []string) {
		fmt.Println("Hello world :)")
	},
}
