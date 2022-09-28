import 'package:flutter/material.dart';
import 'package:flutter_smart_dialog/flutter_smart_dialog.dart';
import 'package:window_size/window_size.dart';
import 'package:flutter/foundation.dart';

import 'homepage.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({Key? key}) : super(key: key);

  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    if (defaultTargetPlatform == TargetPlatform.linux ||
        defaultTargetPlatform == TargetPlatform.macOS ||
        defaultTargetPlatform == TargetPlatform.windows) {
      setWindowTitle("Like Trading Pallet Merger");
    }

    return MaterialApp(
      navigatorObservers: [FlutterSmartDialog.observer],
      // here
      builder: FlutterSmartDialog.init(),
      debugShowCheckedModeBanner: false,
      title: 'Pallet Merger',
      theme: ThemeData(

        primarySwatch: Colors.blue,
      ),
      home:  HomePage(),
    );
  }
}
