\name{xl.get.excel}
\alias{xl.get.excel}

\title{
Function which return reference to Excel application
}
\description{
Returns reference to Microsoft Excel application. If there is no instance already in existence it will create a new instance. 
}
\usage{
xl.get.excel()
}
\arguments{

}
%- maybe also 'usage' for other objects documented here.
\details{
%%  ~~ If necessary, more details than the description above ~~
}
\value{
An object of class 'COMIDispatch' (as returned by 'COMCreate' from RDCOMClient package).

%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
Gregory Demin <excel.link.feedback@gmail.com>
}


\examples{

\dontrun{
xls=xl.get.excel()

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }