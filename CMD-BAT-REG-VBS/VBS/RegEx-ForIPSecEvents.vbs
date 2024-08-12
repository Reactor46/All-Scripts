IKE security association[ -:](?<IKEAssoc>.*)|
\sMode[ -:].*\r\n(?<IKEMode>.*)|
Kerberos based Identity[ -:](?<KerbID>.*)|
Peer SHA Thumbprint[ -:](?<CertID>.*)|
Preshared[ -:](?<PreSharedID>.*)|
Source IP Address[ -:](?<SourceIP>.*)
Source IP Address Mask[ -:](?<SourceIPMask>.*)
Destination IP Address[ -:](?<DestIP>.*)
Destination IP Address Mask[ -:](?<DestIPMask>.*)
Protocol[ -:](?<Proto>.*)
Source Port[ -:](?<SourcePort>.*)
Destination Port[ -:](?<DestPort>.*)|
ESP Algorithm[ -:](?<ESPAlg>.*)
HMAC Algorithm[ -:](?<HMACAlg>.*)
AH Algorithm[ -:](?<AHAlg>.*)|
InboundSpi[ -:](?<InSPI>.*)
OutboundSpi[ -:](?<OutSPI>.*)|
IKE Local Addr[ -:](?<LocalAddr>.*)
IKE Peer Addr[ -:](?<PeerAddr>.*)
|
\s*Failure Point[ -:].*\r\n(?<FailurePoint>.*)
\s*Failure Reason[ -:].*\r\n(?<FailureReason>.*)
