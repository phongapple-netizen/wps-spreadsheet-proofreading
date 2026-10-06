package main

import (
	"errors"
	"os"
	"path/filepath"
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
			}, func() error { return step("health") }, func() error { return step("run") }, func(stop func() error, registrationCompleted bool) error {
				if registrationCompleted != (fail == "run") {
					t.Fatalf("incorrect registration completion state: %v", registrationCompleted)
				}
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
	}, func() error { return nil }, func() error { t.Fatal("must not write Run after registration failure"); return nil }, func(stop func() error, registrationCompleted bool) error {
		if registrationCompleted {
			t.Fatal("failed registration must not be completed")
		}
		if stop == nil {
			t.Fatal("must retain service cleanup handle")
		}
		return errors.New("rollback")
	})
	if err == nil {
		t.Fatal("must report failed install")
	}
}

func TestFailedUpgradeRegistrationPreservesExistingPlugin(t *testing.T) {
	for _, failure := range []string{"backup", "replace"} {
		t.Run(failure, func(t *testing.T) {
			t.Setenv("APPDATA", t.TempDir())
			filename, err := publishPath()
			if err != nil {
				t.Fatal(err)
			}
			if err := os.MkdirAll(filepath.Dir(filename), 0o700); err != nil {
				t.Fatal(err)
			}
			original := `<?xml version="1.0"?><jsplugins><jspluginonline name="wps-spreadsheet-proofreading" type="et" url="http://127.0.0.1:3892/" debug="old"/><jspluginonline name="wps-text-proofreading" type="wps"/><jspluginonline name="other-addon"/></jsplugins>`
			if err := os.WriteFile(filename, []byte(original), 0o600); err != nil {
				t.Fatal(err)
			}
			if failure == "backup" {
				// A directory at the backup destination forces a real backup error.
				if err := os.Mkdir(filename+".wps-spreadsheet-proofreading.bak", 0o700); err != nil {
					t.Fatal(err)
				}
			} else {
				// Windows does not grant delete sharing to os.Open. Holding the
				// old file open permits reads/backup but prevents its replacement.
				locked, err := os.Open(filename)
				if err != nil {
					t.Fatal(err)
				}
				defer locked.Close()
			}
			stopped, runCalled := false, false
			oldRun := stringRunValue(`"C:\Existing\WPSSpreadsheetProofreadingServer.exe" --serve`)
			currentRun := oldRun
			err = installSteps(func() error { return register(true) }, func() (func() error, error) {
				return func() error { stopped = true; return nil }, nil
			}, func() error { return nil }, func() error {
				runCalled = true
				currentRun = stringRunValue("new service")
				return nil
			}, func(stop func() error, registrationCompleted bool) error {
				if registrationCompleted {
					t.Error("failed registration must not trigger publish.xml cleanup")
					if err := register(false); err != nil {
						t.Error(err)
					}
				}
				if stop == nil {
					t.Fatal("missing service cleanup")
				}
				return stop()
			})
			if err == nil {
				t.Fatal("expected real registration failure")
			}
			data, readErr := os.ReadFile(filename)
			if readErr != nil || string(data) != original {
				t.Fatalf("old plugin registration changed: %s err=%v", data, readErr)
			}
			if !stopped || runCalled || currentRun != oldRun {
				t.Fatalf("stopped=%v runCalled=%v RunChanged=%v", stopped, runCalled, currentRun != oldRun)
			}
		})
	}
}
