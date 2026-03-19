import React from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, 
  PieChart, Pie, Cell, Legend, LineChart, Line, AreaChart, Area
} from 'recharts';

const COLORS = ['#006E62', '#00A389', '#FFBB28', '#FF8042', '#0088FE', '#00C49F', '#FFBB28', '#FF8042'];

interface ChartContainerProps {
  title: string;
  children: React.ReactNode;
}

export const ChartContainer: React.FC<ChartContainerProps> = ({ title, children }) => (
  <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
    <h3 className="text-lg font-semibold text-gray-800 mb-6">{title}</h3>
    <div className="h-[300px] w-full">
      <ResponsiveContainer width="100%" height="100%">
        {children as React.ReactElement}
      </ResponsiveContainer>
    </div>
  </div>
);

interface SimpleBarChartProps {
  data: { name: string; value: number }[];
  dataKey?: string;
}

export const SimpleBarChart: React.FC<SimpleBarChartProps> = ({ data, dataKey = "value" }) => (
  <BarChart data={data}>
    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f0f0f0" />
    <XAxis dataKey="name" fontSize={12} tickLine={false} axisLine={false} />
    <YAxis fontSize={12} tickLine={false} axisLine={false} />
    <Tooltip 
      contentStyle={{ backgroundColor: '#fff', borderRadius: '8px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }}
    />
    <Bar dataKey={dataKey} fill="#006E62" radius={[4, 4, 0, 0]} />
  </BarChart>
);

interface SimplePieChartProps {
  data: { name: string; value: number }[];
}

export const SimplePieChart: React.FC<SimplePieChartProps> = ({ data }) => (
  <PieChart>
    <Pie
      data={data}
      cx="50%"
      cy="50%"
      innerRadius={60}
      outerRadius={80}
      paddingAngle={5}
      dataKey="value"
    >
      {data.map((entry, index) => (
        <Cell key={`cell-${index}`} fill={COLORS[index % COLORS.length]} />
      ))}
    </Pie>
    <Tooltip />
    <Legend verticalAlign="bottom" height={36}/>
  </PieChart>
);
