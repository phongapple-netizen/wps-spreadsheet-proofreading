package main

import (
	"errors"
	"reflect"
	"testing"
)

func TestWPSRunningCSV(t *testing.T) {
	for _, name := range []string{"wps.exe", "ET.EXE", "wpp.exe", "explorer.exe"} {
		running, err := wpsRunningCSV(`"` + name + `","42","Console","1","1,000 K"`)
		if err != nil || running != (name != "explorer.exe") {
			t.Fatalf("%s: running=%v err=%v", name, running, err)
		}
	}
	for _, output := range []string{"", "Access denied", `"wps.exe","42"`, `"broken`} {
		if _, err := wpsRunningCSV(output); err == nil {
			t.Fatalf("must reject unknown process state: %q", output)
		}
	}
}

func TestInstallReadinessOrderAndRollback(t *testing.T) {
	for _, fail := range []string{"", "start", "health", "register", "run"} {
		t.Run(fail, func(t *testing.T) {
			var calls []string
			step := func(name string) error {
				calls = append(calls, name)
				if name == fail {
					return errors.New(name)
				}
				return nil
			}
			err := installSteps(func() error { return step("register") }, func() (func() error, error) {
				if err := step("start"); err != nil {
					return nil, err
				}
				return func() error { return step("stop") }, nil
			}, func() error { return step("health") }, func() error { return step("run") }, func(stop func() error) error {
				calls = append(calls, "rollback")
				if stop != nil {
					return stop()
				}
				return nil
			})
			want := []string{"start", "health", "register", "run"}
			if fail != "" {
				for i, name := range want {
					if name == fail {
						want = want[:i+1]
						break
					}
				}
				want = append(want, "rollback")
				if fail != "start" {
					want = append(want, "stop")
				}
			}
			if (err != nil) != (fail != "") || !reflect.DeepEqual(calls, want) {
				t.Fatalf("calls=%v want=%v err=%v", calls, want, err)
			}
		})
	}
}

func TestInstallRegistrationFailureReportsRollbackFailure(t *testing.T) {
	err := installSteps(func() error { return errors.New("registration") }, func() (func() error, error) {
		return func() error { return nil }, nil
	}, func() error { return nil }, func() error { t.Fatal("must not write Run after registration failure"); return nil }, func(stop func() error) error {
		if stop == nil {
			t.Fatal("must retain service cleanup handle")
		}
		return errors.New("rollback")
	})
	if err == nil {
		t.Fatal("must report failed install")
	}
}
