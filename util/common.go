package util

import (
	"path"
	"strings"
)

// GetFilename get file name by filepath
func GetFilename(filepath string) string {
	var filenameWithSuffix string
	filenameWithSuffix = path.Base(filepath)
	var fileSuffix string
	fileSuffix = path.Ext(filenameWithSuffix)

	var filenameOnly string
	filenameOnly = strings.TrimSuffix(filenameWithSuffix, fileSuffix)
	return filenameOnly
}
