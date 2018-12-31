RunAction "Initialize", "2 - 2"
RunAction "Setup", allIterations
RunAction "Good1", allIterations
RunAction "GrdCfgMgr_DataPurgeWithSPTNumbers [GrdCfgMgr_PurgeData] [2]", oneIteration
RunAction "Good2", allIterations
RunAction "Reject1", allIterations
RunAction "Reject2", allIterations
RunAction "Fault1", allIterations
RunAction "Fault2", allIterations
RunAction "CleanUp", allIterations
