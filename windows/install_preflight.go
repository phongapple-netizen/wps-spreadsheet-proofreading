package main

import (
	"context"
	"encoding/csv"
	"fmt"
	"os/exec"
	"strings"
	"time"
)

func wpsRunningCSV(output string) (bool, error) {
	records, err := csv.NewReader(strings.NewReader(output)).ReadAll()
	if err != nil || len(records) == 0 {
		return false, fmt.Errorf("无法读取进程列表")
	}
	for _, record := range records {
		if len(record) != 5 {
			return false, fmt.Errorf("进程列表格式异常")
		}
		switch strings.ToLower(record[0]) {
		case "wps.exe", "et.exe", "wpp.exe":
			return true, nil
		}
	}
	return false, nil
}

func requireWPSClosed() error {
	ctx, cancel := context.WithTimeout(context.Background(), 5*time.Second)
	defer cancel()
	cmd := exec.CommandContext(ctx, "tasklist.exe", "/FO", "CSV", "/NH")
	hideWindow(cmd)
	output, err := cmd.Output()
	if err != nil {
		return fmt.Errorf("无法确认 WPS 是否已退出，安装未继续: %w", err)
	}
	running, err := wpsRunningCSV(string(output))
	if err != nil {
		return err
	}
	if running {
		return fmt.Errorf("请保存所有文档并完全退出 WPS（含托盘和后台进程）后重试；安装未继续")
	}
	return nil
}
