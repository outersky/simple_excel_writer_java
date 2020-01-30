# rm -rf /tmp/trace && mkdir /tmp/trace && java -agentlib:native-image-agent=config-merge-dir=/tmp/trace -jar build/libs/*-0.1-all.jar
native-image --verbose --no-server --no-fallback --report-unsupported-elements-at-runtime --allow-incomplete-classpath -H:+TraceServiceLoaderFeature -H:+TraceClassInitialization  -H:+ReportExceptionStackTraces -cp target/*-1.0.jar
