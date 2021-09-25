package mixoder

import (
	"fmt"
	"sync"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/golang/protobuf/proto"
	"go.uber.org/zap"
)

type ProcMon struct {
	stopChannel chan bool
	logger      *zap.SugaredLogger

	monitor     []string
	monitorLock sync.RWMutex

	procEventConsumers []chan ProcEvent

	// comPort  string
	// baudRate uint

	// deej   *Deej

	// stopChannel chan bool
	// connected   bool
	// connOptions serial.OpenOptions
	// conn        io.ReadWriteCloser

	// lastKnownNumSliders        int
	// currentSliderPercentValues []float32

	// sliderMoveConsumers []chan SliderMoveEvent
}

type Process struct {
	PID  int
	Name string
}

type ProcEvent struct {
	PID    int
	Name   string
	Action ProcAction
}

type ProcAction int32

const (
	PA_START ProcAction = iota
	PA_STOP
	PA_UNKNOWN
)

var ProcActionName = map[int32]string{
	0: "START",
	1: "STOP",
	2: "UNKNOWN",
}

var ProcActionValue = map[string]int32{
	"START":   0,
	"STOP":    1,
	"UNKNOWN": 2,
}

func (pa ProcAction) String() string {
	return proto.EnumName(ProcActionName, int32(pa))
}

// NewProcMon returns a new ProcMon instance
func NewProcMon(logger *zap.SugaredLogger, monitorProcs []string) (*ProcMon, error) {
	pm := &ProcMon{
		logger:             logger.Named("proc_mon"),
		stopChannel:        make(chan bool),
		running:            []Process{},
		runningLock:        &sync.RWMutex{},
		procEventConsumers: []chan SliderMoveEvent{},
		monitor:            monitorProcs,
		monitorLock:        &sync.RWMutex{},
	}

	pm.logger.Debug("Created ProcMon instance")

	return pm, nil
}

func (pm *ProcMon) start() {
	responseChan := make(chan ProcEvent)
	stopChannel := make(chan bool)

	go WatchForNewProcesses(responseChan, stopChannel)

	select {
	case <-pm.stopChannel:
		stopChannel <- true
	case <-responseChan:
		go filterProcessEvents()
	}
}

func (pm *ProcMon) setMonitorProcs(processes []string) {
	pm.monitorLock.Lock()
	defer pm.monitorLock.Unlock()
	pm.monitor = processes
}

func (pm *ProcMon) filterProcessEvent(e ProcEvent) {
	pm.runningLock.Lock()
	defer pm.runningLock.Unlock()

}

// SubscribeToProcEvents returns an unbuffered channel that receives
// a ProcEvent struct every time a process is started or destroyed
func (pm *ProcMon) SubscribeToProcEvents() chan ProcEvent {
	ch := make(chan ProcEvent)
	pm.procEventConsumers = append(pm.procEventConsumers, ch)

	return ch
}

func (pm *ProcMon) WatchForNewProcesses(responseChan chan ProcEvent, stopChan chan bool) {
	method := "Win32_ProcessTrace"
	// init COM, oh yeah
	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, _ := oleutil.CreateObject("WbemScripting.SWbemLocator")
	defer unknown.Release()

	wmi, _ := unknown.QueryInterface(ole.IID_IDispatch)
	defer wmi.Release()

	// service is a SWbemServices
	serviceRaw, _ := oleutil.CallMethod(wmi, "ConnectServer")
	service := serviceRaw.ToIDispatch()
	defer service.Release()

	// result is a SWBemObjectSet
	resultRaw, err := oleutil.CallMethod(service, "ExecNotificationQuery", fmt.Sprintf("SELECT * FROM %s", method))
	if err != nil {
		m.logger.Errorf("error querying WMI: %v", err)
		return
	}
	result := resultRaw.ToIDispatch()
	defer result.Release()

	// m.deej.config.SliderMapping.iterate(func(sliderIdx int, targets []string) {
	// 	for _, target := range targets {
	// 		m.logger.Debugf("Looking for these targets: %v", target)
	// 	}
	// })

	for {
		select {
		case <-stopChan:
			pm.logger.Debugf("stopping watch for processes")
			return
		default:
			eventRaw, err := oleutil.CallMethod(result, "NextEvent")
			if err != nil {
				pm.logger.Errorf("error getting NextEvent from notification query: %v", err)
				return
			}
			event := eventRaw.ToIDispatch()
			defer event.Release()

			name, err := oleutil.GetProperty(event, "ProcessName")
			if err != nil {
				pm.logger.Errorf("error getting process name from process event: %v", err)
			}
			pid, err := oleutil.GetProperty(event, "ProcessID")
			if err != nil {
				pm.logger.Errorf("error getting process id from process event: %v", err)
			}

			pm.logger.Debugf("%v: %v", name.ToString(), pid.Value())

			procEvent := &ProcEvent{
				Name:   name,
				PID:    pid,
				Action: PA_UNKNOWN,
			}

			responseChan <- procEvent

		}

		// m.deej.config.SliderMapping.iterate(func(sliderIdx int, targets []string) {
		// 	for _, target := range targets {

		// 		// ignore special transforms
		// 		if m.targetHasSpecialTransform(target) {
		// 			continue
		// 		}

		// 		// safe to assume this has a single element because we made sure there's no special transform
		// 		target = m.resolveTarget(target)[0]

		// 		if target == strings.ToLower(name.ToString()) {
		// 			m.logger.Debugf("found match %v: %v = %v", name.ToString(), pid.Value(), target)
		// 			m.refreshSessions(false)
		// 			m.logger.Debugf("sessions: %+v", m.m)
		// 			break
		// 		}
		// 	}
		// })

	}

}
