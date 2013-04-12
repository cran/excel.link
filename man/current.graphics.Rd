\name{current.graphics}
\alias{current.graphics}

\title{
Auxilary function for export graphics to Microsoft Excel 
}
\description{
Save graphics from windows plotting device to temporary file and return path to this file
}
\usage{
current.graphics(type="emf",\dots)
}
\arguments{
\item{type}{file type}
\item{\dots}{arguments for internally used savePlot function}
}
%- maybe also 'usage' for other objects documented here.

\value{
Path to file with saved graphics with class attribute 'current.graphics' and attribute 'temp.file'=TRUE

%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}

\author{
Gregory Demin <excel.link.feedback@gmail.com>
}


\examples{

\dontrun{
xl.workbook.add()
plot(sin)
xl[a1]=current.graphics()
plot(cos)
cos.plot=current.graphics()
xl.sheet.add()
xl[a1]=list("Cosine plotting",cos.plot,"End of cosine plotting")

}
  
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
% \keyword{ ~xl }
\keyword{ utilities }