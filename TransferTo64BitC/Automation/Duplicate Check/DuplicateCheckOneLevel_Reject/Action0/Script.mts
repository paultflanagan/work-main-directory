﻿RunAction "Initialize", "2 - 2"
RunAction "Setup", allIterations
RunAction "StartLot", oneIteration
RunAction "Good1", allIterations
RunAction "EndLot_QuarantineReject", oneIteration
RunAction "VerifyEmail_App_GoodLot [VerifyEmail_App_GoodLot]", oneIteration
RunAction "Good2", allIterations
RunAction "Reject1", allIterations
RunAction "Reject2", allIterations
RunAction "Fault1", allIterations
RunAction "Fault2", allIterations
RunAction "CleanUp", allIterations